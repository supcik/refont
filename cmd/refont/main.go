package main

import (
	"archive/zip"
	"flag"
	"fmt"
	"io"
	"os"
	"path"

	"github.com/beevik/etree"
)

const (
	style = iota
	sharedStrings
	workbook
	worksheet
)

type context struct {
	fromFont    *string
	toFont      *string
	doUnprotect *bool
}

func patchXML(dst io.Writer, src io.ReadCloser, patchType int, ctx *context) error {
	doc := etree.NewDocument()
	if _, err := doc.ReadFrom(src); err != nil {
		return err
	}
	switch {
	case patchType == style:
		for _, f := range doc.FindElements("//fonts/font/name") {
			if f.SelectAttr("val").Value == *ctx.fromFont {
				f.SelectAttr("val").Value = *ctx.toFont
			}
		}
	case patchType == sharedStrings:
		for _, f := range doc.FindElements("//rFont") {
			if f.SelectAttr("val").Value == *ctx.fromFont {
				f.SelectAttr("val").Value = *ctx.toFont
			}
		}

	case patchType == workbook:
		for _, f := range doc.FindElements("//workbookProtection") {
			f.Parent().RemoveChild(f)
		}
		for _, f := range doc.FindElements("//fileSharing") {
			f.Parent().RemoveChild(f)
		}

	case patchType == worksheet:
		for _, f := range doc.FindElements("//sheetProtection") {
			f.Parent().RemoveChild(f)
		}
	}

	doc.WriteTo(dst)
	return nil
}

func copy(w *zip.Writer, f *zip.File, ctx *context) error {
	dst, err := w.Create(f.Name)
	if err != nil {
		return err
	}
	src, err := f.Open()
	if err != nil {
		return err
	}
	defer src.Close()

	switch {
	case f.Name == "xl/styles.xml":
		patchXML(dst, src, style, ctx)
	case f.Name == "xl/sharedStrings.xml":
		patchXML(dst, src, sharedStrings, ctx)
	case *ctx.doUnprotect && f.Name == "xl/workbook.xml":
		patchXML(dst, src, workbook, ctx)
	case *ctx.doUnprotect && path.Dir(f.Name) == "xl/worksheets":
		patchXML(dst, src, worksheet, ctx)
	default:
		io.Copy(dst, src)
	}

	return nil
}

func process(fileName string, ctx *context) {
	zipReader, err := zip.OpenReader(fileName)
	if err != nil {
		panic(err)
	}
	defer zipReader.Close()

	zipFileWriter, err := os.CreateTemp(path.Dir(fileName), "tmp-*.xlsx")
	if err != nil {
		panic(err)
	}

	zipWriter := zip.NewWriter(zipFileWriter)
	for _, src := range zipReader.File {
		err = copy(zipWriter, src, ctx)
		if err != nil {
			panic(err)
		}
	}

	zipWriter.Close()
	zipFileWriter.Close()
	os.Rename(zipFileWriter.Name(), fileName)
	os.Remove(zipFileWriter.Name())
}

func main() {

	flag.Usage = func() {
		fmt.Printf("Usage: %s [OPTIONS] files ...\n", path.Base(os.Args[0]))
		flag.PrintDefaults()
	}

	var ctx context
	ctx.fromFont = flag.String("from", "Geneva", "From Font")
	ctx.toFont = flag.String("to", "Arial", "To Font")
	ctx.doUnprotect = flag.Bool("unprotect", false, "Unprotect Workbook")
	flag.Parse()
	for _, fileName := range flag.Args() {
		process(fileName, &ctx)
	}
}
