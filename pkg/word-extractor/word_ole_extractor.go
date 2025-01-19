package word_extractor

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strings"
	"unicode/utf8"

	"github.com/richardlehane/mscfb"
)

type unbufferedReaderAt struct {
	R io.ReadSeeker
	N int64
}

func NewUnbufferedReaderAt(r io.ReadSeeker) io.ReaderAt {
	return &unbufferedReaderAt{R: r}
}

func (u *unbufferedReaderAt) ReadAt(p []byte, off int64) (n int, err error) {
	if off < u.N {
		return 0, errors.New("invalid offset")
	}
	diff := off - u.N
	written, err := io.CopyN(ioutil.Discard, u.R, diff)
	u.N += written
	if err != nil {
		return 0, err
	}

	n, err = u.R.Read(p)
	u.N += int64(n)
	return
}

const (
	sprmCFRMarkDel = 0x00
)

// WordOleExtractor handles extraction of text from OLE-based Word files
type WordOleExtractor struct {
	pieces        []Piece
	bookmarks     map[string]Bookmark
	boundaries    Boundaries
	taggedHeaders []TaggedHeader
}

type Piece struct {
	StartCp      int
	StartStream  int
	TotLength    int
	StartFilePos int
	Unicode      bool
	Bpc          int
	Size         int
	Text         string
	Length       int
	EndCp        int
	EndStream    int
	EndFilePos   int
}

type Bookmark struct {
	Start int
	End   int
}

type Boundaries struct {
	FcMin      int
	CcpText    int
	CcpFtn     int
	CcpHdd     int
	CcpAtn     int
	CcpEdn     int
	CcpTxbx    int
	CcpHdrTxbx int
}

type TaggedHeader struct {
	Type string
	Text string
}

// NewWordOleExtractor creates a new WordOleExtractor instance
func NewWordOleExtractor() *WordOleExtractor {
	return &WordOleExtractor{
		pieces:        make([]Piece, 0),
		bookmarks:     make(map[string]Bookmark),
		taggedHeaders: make([]TaggedHeader, 0),
	}
}

// Extract implements the DocumentExtractor interface
func (w *WordOleExtractor) Extract(reader io.ReadSeeker) (*Document, error) {

	buffer, err := readStream(reader, "WordDocument")
	if err != nil {
		return nil, err
	}

	return w.extractWordDocument(reader, buffer)
}

func (w *WordOleExtractor) extractWordDocument(reader io.ReadSeeker, buffer []byte) (*Document, error) {
	// Check magic number (0xA5EC)
	magic := binary.LittleEndian.Uint16(buffer[0:2])
	if magic != 0xA5EC {
		return nil, errors.New("invalid Word document: incorrect magic number")
	}

	// Get flags and determine table stream name
	flags := binary.LittleEndian.Uint16(buffer[0x0A:0x0C])
	streamName := "1Table"
	if (flags & 0x0200) == 0 {
		streamName = "0Table"
	}

	tableBuffer, err := readStream(reader, streamName)
	if err != nil {
		return nil, err
	}

	// Extract document boundaries
	w.boundaries = Boundaries{
		FcMin:      int(binary.LittleEndian.Uint32(buffer[0x0018:0x001C])),
		CcpText:    int(binary.LittleEndian.Uint32(buffer[0x004C:0x0050])),
		CcpFtn:     int(binary.LittleEndian.Uint32(buffer[0x0050:0x0054])),
		CcpHdd:     int(binary.LittleEndian.Uint32(buffer[0x0054:0x0058])),
		CcpAtn:     int(binary.LittleEndian.Uint32(buffer[0x005C:0x0060])),
		CcpEdn:     int(binary.LittleEndian.Uint32(buffer[0x0060:0x0064])),
		CcpTxbx:    int(binary.LittleEndian.Uint32(buffer[0x0064:0x0068])),
		CcpHdrTxbx: int(binary.LittleEndian.Uint32(buffer[0x0068:0x006C])),
	}

	// Extract document components
	if err := w.writeBookmarks(buffer, tableBuffer); err != nil {
		return nil, err
	}
	if err := w.writePieces(buffer, tableBuffer); err != nil {
		return nil, err
	}
	if err := w.writeCharacterProperties(buffer, tableBuffer); err != nil {
		return nil, err
	}
	if err := w.writeParagraphProperties(buffer, tableBuffer); err != nil {
		return nil, err
	}
	if err := w.normalizeHeaders(buffer, tableBuffer); err != nil {
		return nil, err
	}

	return w.buildDocument()
}

func (w *WordOleExtractor) buildDocument() (*Document, error) {
	doc := NewDocument()
	start := 0

	// Extract body text
	doc.Body = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpText))
	start += w.boundaries.CcpText

	// Extract footnotes if present
	if w.boundaries.CcpFtn > 0 {
		doc.Footnotes = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpFtn-1))
		start += w.boundaries.CcpFtn
	}

	// Extract headers and footers if present
	if w.boundaries.CcpHdd > 0 {
		headers := make([]string, 0)
		footers := make([]string, 0)
		for _, header := range w.taggedHeaders {
			if header.Type == "headers" {
				headers = append(headers, header.Text)
			} else if header.Type == "footers" {
				footers = append(footers, header.Text)
			}
		}
		doc.Headers = cleanText(join(headers, ""))
		doc.Footers = cleanText(join(footers, ""))
		start += w.boundaries.CcpHdd
	}

	// Extract annotations if present
	if w.boundaries.CcpAtn > 0 {
		doc.Annotations = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpAtn-1))
		start += w.boundaries.CcpAtn
	}

	// Extract endnotes if present
	if w.boundaries.CcpEdn > 0 {
		doc.Endnotes = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpEdn-1))
		start += w.boundaries.CcpEdn
	}

	// Extract textboxes if present
	if w.boundaries.CcpTxbx > 0 {
		doc.Textboxes = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpTxbx-1))
		start += w.boundaries.CcpTxbx
	}

	// Extract header textboxes if present
	if w.boundaries.CcpHdrTxbx > 0 {
		doc.HeaderTextboxes = cleanText(w.getTextRangeByCP(start, start+w.boundaries.CcpHdrTxbx-1))
		start += w.boundaries.CcpHdrTxbx
	}

	return doc, nil
}

// Helper functions

func readStream(reader io.ReadSeeker, name string) ([]byte, error) {
	var readerAt io.ReaderAt
	if _, ok := reader.(*os.File); !ok {
		readerAt = NewUnbufferedReaderAt(reader)
	} else {
		readerAt = reader.(io.ReaderAt)
	}

	readerAt.(io.ReadSeeker).Seek(0, io.SeekStart)
	cfb, err := mscfb.New(readerAt)
	if err != nil {
		return nil, err
	}

	for entry, err := cfb.Next(); err == nil; entry, err = cfb.Next() {
		if entry.Name == name {
			buf := new(bytes.Buffer)
			_, err := buf.ReadFrom(cfb)
			if err != nil {
				return nil, err
			}
			return buf.Bytes(), nil
		}
	}
	return nil, errors.New("stream not found: " + name)
}

// Helper functions for text manipulation and processing would go here

func (w *WordOleExtractor) writeBookmarks(buffer, tableBuffer []byte) error {
	fcSttbfBkmk := binary.LittleEndian.Uint32(buffer[0x0142:0x0146])
	lcbSttbfBkmk := binary.LittleEndian.Uint32(buffer[0x0146:0x014A])
	fcPlcfBkf := binary.LittleEndian.Uint32(buffer[0x014A:0x014E])
	lcbPlcfBkf := binary.LittleEndian.Uint32(buffer[0x014E:0x0152])
	fcPlcfBkl := binary.LittleEndian.Uint32(buffer[0x0152:0x0156])
	lcbPlcfBkl := binary.LittleEndian.Uint32(buffer[0x0156:0x015A])

	if lcbSttbfBkmk == 0 {
		return nil
	}

	sttbfBkmk := tableBuffer[fcSttbfBkmk : fcSttbfBkmk+lcbSttbfBkmk]
	plcfBkf := tableBuffer[fcPlcfBkf : fcPlcfBkf+lcbPlcfBkf]
	plcfBkl := tableBuffer[fcPlcfBkl : fcPlcfBkl+lcbPlcfBkl]

	fcExtend := binary.LittleEndian.Uint16(sttbfBkmk[0:2])
	if fcExtend != 0xFFFF {
		return errors.New("unexpected single-byte bookmark data")
	}

	offset := 6 // Skip header bytes
	index := 0

	for offset < int(lcbSttbfBkmk) {
		length := int(binary.LittleEndian.Uint16(sttbfBkmk[offset:offset+2])) * 2
		segment := sttbfBkmk[offset+2 : offset+2+length]
		cpStart := binary.LittleEndian.Uint32(plcfBkf[index*4 : index*4+4])
		cpEnd := binary.LittleEndian.Uint32(plcfBkl[index*4 : index*4+4])
		w.bookmarks[string(segment)] = Bookmark{Start: int(cpStart), End: int(cpEnd)}
		offset += length + 2
	}

	return nil
}

func (w *WordOleExtractor) writePieces(buffer, tableBuffer []byte) error {
	pos := binary.LittleEndian.Uint32(buffer[0x01A2:0x01A6])

	// Skip initial flags
	for tableBuffer[pos] == 1 {
		skip := binary.LittleEndian.Uint16(tableBuffer[pos+1:])
		pos += 3 + uint32(skip)
	}

	if tableBuffer[pos] != 2 {
		return errors.New("corrupted Word file")
	}
	pos++

	pieceTableSize := binary.LittleEndian.Uint32(tableBuffer[pos:])
	pos += 4

	pieces := (pieceTableSize - 4) / 12
	var startCp, startStream int

	for x := uint32(0); x < pieces; x++ {
		offset := pos + ((pieces + 1) * 4) + (x * 8) + 2
		startFilePos := binary.LittleEndian.Uint32(tableBuffer[offset:])

		unicode := true
		if (startFilePos & 0x40000000) != 0 {
			unicode = false
			startFilePos &= ^uint32(0x40000000)
			startFilePos = startFilePos / 2
		}

		lStart := binary.LittleEndian.Uint32(tableBuffer[pos+x*4:])
		lEnd := binary.LittleEndian.Uint32(tableBuffer[pos+(x+1)*4:])
		totLength := lEnd - lStart

		piece := Piece{
			StartCp:      startCp,
			StartStream:  startStream,
			TotLength:    int(totLength),
			StartFilePos: int(startFilePos),
			Unicode:      unicode,
			Bpc:          1,
		}

		if unicode {
			piece.Bpc = 2
		}

		piece.Size = piece.Bpc * int(lEnd-lStart)

		// Extract text correctly based on unicode flag
		textBuffer := buffer[startFilePos : startFilePos+uint32(piece.Size)]
		if unicode {
			text, err := bufferToUCS2String(textBuffer)
			if err != nil {
				fmt.Printf("Error converting text buffer: %v\n", err)
			}
			piece.Text = text
		} else {
			// For non-unicode text, convert each byte using binaryToUnicode mapping
			var text strings.Builder
			for _, b := range textBuffer {
				text.WriteString(string(b))
			}
			piece.Text = binaryToUnicode(text.String())
		}

		// Update Length to use rune count
		piece.Length = utf8.RuneCountInString(piece.Text)
		piece.EndCp = piece.StartCp + piece.Length
		piece.EndStream = piece.StartStream + piece.Size
		piece.EndFilePos = piece.StartFilePos + piece.Size

		startCp = piece.EndCp
		startStream = piece.EndStream

		w.pieces = append(w.pieces, piece)
	}

	return nil
}

func (w *WordOleExtractor) normalizeHeaders(buffer, tableBuffer []byte) error {
	fcPlcfhdd := binary.LittleEndian.Uint32(buffer[0x00F2:0x00F6])
	lcbPlcfhdd := binary.LittleEndian.Uint32(buffer[0x00F6:0x00FA])

	if lcbPlcfhdd < 8 {
		return nil
	}

	offset := w.boundaries.CcpText + w.boundaries.CcpFtn
	ccpHdd := w.boundaries.CcpHdd

	plcHdd := tableBuffer[fcPlcfhdd : fcPlcfhdd+lcbPlcfhdd]
	plcHddCount := lcbPlcfhdd / 4

	start := offset + int(binary.LittleEndian.Uint32(plcHdd[0:4]))

	for i := uint32(1); i < plcHddCount; i++ {
		end := offset + int(binary.LittleEndian.Uint32(plcHdd[i*4:]))
		if end > offset+ccpHdd {
			end = offset + ccpHdd
		}

		text := w.getTextRangeByCP(start, end)
		story := int(i - 1)

		header := TaggedHeader{Text: text}

		switch {
		case story < 3:
			header.Type = "footnoteSeparators"
		case story < 6:
			header.Type = "endSeparators"
		case story%6 == 0 || story%6 == 1 || story%6 == 4:
			header.Type = "headers"
		case story%6 == 2 || story%6 == 3 || story%6 == 5:
			header.Type = "footers"
		}

		w.taggedHeaders = append(w.taggedHeaders, header)

		if !containsNonWhitespace(text) {
			w.replaceSelectedRange(start, end, "\x00")
		} else {
			w.replaceSelectedRange(end-1, end, "\x00")
		}

		start = end
	}

	return nil
}

// Helper functions

func getPieceIndexByCP(pieces []Piece, position int) int {
	for i, piece := range pieces {
		if position <= piece.EndCp {
			return i
		}
	}
	return len(pieces) - 1
}

func (w *WordOleExtractor) getTextRangeByCP(start, end int) string {
	startPiece := getPieceIndexByCP(w.pieces, start)
	endPiece := getPieceIndexByCP(w.pieces, end)

	var result []string
	for i := startPiece; i <= endPiece; i++ {
		piece := w.pieces[i]
		runes := []rune(piece.Text)
		xstart := 0
		if i == startPiece {
			xstart = start - piece.StartCp
		}
		xend := piece.Length
		if i == endPiece {
			xend = end - piece.StartCp
		}

		// Handle bounds checking
		if xend > len(runes) {
			xend = len(runes)
		}
		if xstart < 0 {
			xstart = 0
		}
		if xstart < len(runes) {
			result = append(result, string(runes[xstart:xend]))
		}
	}
	return strings.Join(result, "")
}

func getPieceIndexByFilePos(pieces []Piece, position int) int {
	for i, piece := range pieces {
		if position <= piece.EndFilePos {
			return i
		}
	}
	return len(pieces) - 1
}

func fillPieceRange(piece *Piece, start, end int, character string) {
	pieceStart := piece.StartCp
	pieceEnd := pieceStart + piece.Length
	if start < pieceStart {
		start = pieceStart
	}
	if end > pieceEnd {
		end = pieceEnd
	}

	// Convert to rune slice for safe unicode operations
	runes := []rune(piece.Text)
	startIdx := start - pieceStart
	endIdx := end - pieceStart

	// Build the modified text using rune slice operations
	var result []rune
	if startIdx > 0 {
		result = append(result, runes[:startIdx]...)
	}
	result = append(result, []rune(strings.Repeat(character, end-start))...)
	if endIdx < len(runes) {
		result = append(result, runes[endIdx:]...)
	}
	piece.Text = string(result)
}

func fillPieceRangeByFilePos(piece *Piece, start, end int, character string) {
	pieceStart := piece.StartFilePos
	pieceEnd := pieceStart + piece.Size
	if start < pieceStart {
		start = pieceStart
	}
	if end > pieceEnd {
		end = pieceEnd
	}

	// Convert byte offsets to rune indices
	runes := []rune(piece.Text)
	startIdx := (start - pieceStart) / piece.Bpc
	endIdx := (end - pieceStart) / piece.Bpc

	// Build the modified text using rune slice operations
	var result []rune
	if startIdx > 0 {
		result = append(result, runes[:startIdx]...)
	}
	result = append(result, []rune(strings.Repeat(character, (end-start)/piece.Bpc))...)
	if endIdx < len(runes) {
		result = append(result, runes[endIdx:]...)
	}
	piece.Text = string(result)
}

func (w *WordOleExtractor) replaceSelectedRange(start, end int, character string) {
	startPiece := getPieceIndexByCP(w.pieces, start)
	endPiece := getPieceIndexByCP(w.pieces, end)
	for i := startPiece; i <= endPiece; i++ {
		fillPieceRange(&w.pieces[i], start, end, character)
	}
}

func (w *WordOleExtractor) replaceSelectedRangeByFilePos(start, end int, character string) {
	startPiece := getPieceIndexByFilePos(w.pieces, start)
	endPiece := getPieceIndexByFilePos(w.pieces, end)
	for i := startPiece; i <= endPiece; i++ {
		fillPieceRangeByFilePos(&w.pieces[i], start, end, character)
	}
}

func (w *WordOleExtractor) markDeletedRange(start, end int) {
	w.replaceSelectedRangeByFilePos(start, end, "\x00")
}

func processSprms(buffer []byte, offset int, handler func(buffer []byte, offset int, sprm uint16, ispmd uint16, fspec uint8, sgc uint8, spra uint8)) {
	for offset < len(buffer)-1 {
		sprm := binary.LittleEndian.Uint16(buffer[offset:])
		ispmd := sprm & 0x1ff
		fspec := uint8((sprm >> 9) & 0x01)
		sgc := uint8((sprm >> 10) & 0x07)
		spra := uint8(sprm >> 13)

		offset += 2

		handler(buffer, offset, sprm, ispmd, fspec, sgc, spra)

		switch spra {
		case 0, 1:
			offset += 1
		case 2:
			offset += 2
		case 3:
			offset += 4
		case 4, 5:
			offset += 2
		case 6:
			offset += int(buffer[offset]) + 1
		case 7:
			offset += 3
		default:
			return
		}
	}
}

func (w *WordOleExtractor) writeParagraphProperties(buffer, tableBuffer []byte) error {
	fcPlcfbtePapx := binary.LittleEndian.Uint32(buffer[0x0102:0x0106])
	lcbPlcfbtePapx := binary.LittleEndian.Uint32(buffer[0x0106:0x010A])

	plcBtePapxCount := (lcbPlcfbtePapx - 4) / 8
	dataOffset := (plcBtePapxCount + 1) * 4
	plcBtePapx := tableBuffer[fcPlcfbtePapx : fcPlcfbtePapx+lcbPlcfbtePapx]

	for i := uint32(0); i < plcBtePapxCount; i++ {
		papxFkpBlock := binary.LittleEndian.Uint32(plcBtePapx[dataOffset+i*4:])
		papxFkpBlockBuffer := buffer[papxFkpBlock*512 : (papxFkpBlock+1)*512]

		crun := uint32(papxFkpBlockBuffer[511])
		for j := uint32(0); j < crun; j++ {
			rgfc := binary.LittleEndian.Uint32(papxFkpBlockBuffer[j*4:])
			rgfcNext := binary.LittleEndian.Uint32(papxFkpBlockBuffer[(j+1)*4:])

			cbLocation := (uint32(crun)+1)*4 + uint32(j)*13
			cbIndex := uint32(papxFkpBlockBuffer[cbLocation]) * 2

			var grpPrlAndIstd []byte
			if cb := uint32(papxFkpBlockBuffer[cbIndex]); cb != 0 {
				grpPrlAndIstd = papxFkpBlockBuffer[uint32(cbIndex)+1 : uint32(cbIndex)+1+uint32(2*cb)-1]
			} else {
				cb2 := papxFkpBlockBuffer[cbIndex+1]
				grpPrlAndIstd = papxFkpBlockBuffer[uint32(cbIndex)+2 : uint32(cbIndex)+2+uint32(2*cb2)]
			}
			processSprms(grpPrlAndIstd, 2, func(buffer []byte, offset int, sprm uint16, ispmd uint16, fspec uint8, sgc uint8, spra uint8) {
				if sprm == uint16(0x2417) {
					w.replaceSelectedRangeByFilePos(int(rgfc), int(rgfcNext), "\n")
				}
			})
		}
	}
	return nil
}

func (w *WordOleExtractor) writeCharacterProperties(buffer, tableBuffer []byte) error {
	fcPlcfbteChpx := binary.LittleEndian.Uint32(buffer[0x00FA:0x00FE])
	lcbPlcfbteChpx := binary.LittleEndian.Uint32(buffer[0x00FE:0x0102])

	// Skip if no character properties
	if lcbPlcfbteChpx == 0 {
		return nil
	}

	plcBteChpxCount := (lcbPlcfbteChpx - 4) / 8
	dataOffset := (plcBteChpxCount + 1) * 4

	// Validate buffer sizes
	if int(fcPlcfbteChpx+lcbPlcfbteChpx) > len(tableBuffer) {
		return errors.New("invalid table buffer size")
	}

	plcBteChpx := tableBuffer[fcPlcfbteChpx : fcPlcfbteChpx+lcbPlcfbteChpx]
	var lastDeletionEnd int

	for i := uint32(0); i < plcBteChpxCount; i++ {
		binary.LittleEndian.Uint32(plcBteChpx[i*4:])
		chpxFkpBlock := binary.LittleEndian.Uint32(plcBteChpx[dataOffset+i*4:])

		// Calculate block boundaries
		blockStart := chpxFkpBlock * 512
		blockEnd := (chpxFkpBlock + 1) * 512
		if int(blockEnd) > len(buffer) {
			return errors.New("invalid block boundaries")
		}

		chpxFkpBlockBuffer := buffer[blockStart:blockEnd]

		crun := uint32(chpxFkpBlockBuffer[511])

		for j := uint32(0); j < crun; j++ {
			rgfc := binary.LittleEndian.Uint32(chpxFkpBlockBuffer[j*4:])
			rgfcNext := binary.LittleEndian.Uint32(chpxFkpBlockBuffer[(j+1)*4:])
			rgb := uint32(chpxFkpBlockBuffer[(uint32(crun)+1)*4+uint32(j)])
			if rgb == 0 {
				continue
			}

			chpxOffset := rgb * 2

			cb := uint32(chpxFkpBlockBuffer[chpxOffset])
			if int(chpxOffset+1+cb) > len(chpxFkpBlockBuffer) {
				return errors.New("invalid grpprl boundaries")
			}

			grpprl := chpxFkpBlockBuffer[chpxOffset+1 : chpxOffset+1+cb]

			processSprms(grpprl, 0, func(buffer []byte, offset int, sprm uint16, ispmd uint16, fspec uint8, sgc uint8, spra uint8) {
				if ispmd == uint16(sprmCFRMarkDel) {
					if (buffer[offset] & 1) != 1 {
						return
					}

					if lastDeletionEnd == int(rgfc) {
						w.markDeletedRange(lastDeletionEnd, int(rgfcNext))
					} else {
						w.markDeletedRange(int(rgfc), int(rgfcNext))
					}
					lastDeletionEnd = int(rgfcNext)
				}
			})
		}
	}
	return nil
}
