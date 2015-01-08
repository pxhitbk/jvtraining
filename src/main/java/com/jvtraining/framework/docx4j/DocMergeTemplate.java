/*
 * FILENAME
 *     DocMergeTemplate.java
 * FILE LOCATION
 *     $Source$
 * VERSION
 *     $Id$
 *         @version       $Revision$
 *         Check-Out Tag: $Name$
 *         Locked By:     $Lockers$
 * FORMATTING NOTES
 *     * Lines should be limited to 78 characters.
 *     * Files should contain no tabs.
 *     * Indent code using four-character increments.
 * COPYRIGHT
 *     Copyright (C) Vietinbank. Ltd. All rights reserved.
 *     This software is the confidential and proprietary information of
 *     Vietinbank ("Confidential Information"). You shall not
 *     disclose such Confidential Information and shall use it only in
 *     accordance with the terms of the license agreement you entered into
 *     with Vietinbank.
 */

package com.jvtraining.framework.docx4j;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.Calendar;
import java.util.List;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;

// IMPORTS

public class DocMergeTemplate
{
    static String dir;
    static String FONT_SIZE = "20";
    static String GLOBAL_FONT = "Arabic Newspaper";//Times New Roman, Arial
    
    static
    {
        dir = System.getProperty("user.dir") + "/template/";
    }
    static WordprocessingMLPackage wordMLPackage;
    static ObjectFactory factory;

    public static void main(final String[] args) throws Exception
    {
        wordMLPackage = WordprocessingMLPackage.createPackage();
        factory = Context.getWmlObjectFactory();
        Body body = mainDoc().getContents().getBody();
        PageDimensions page = new PageDimensions();

        PgMar pgMar = page.getPgMar();
        pgMar.setLeft(new BigInteger("900"));
        pgMar.setRight(new BigInteger("900"));
        SectPr sectPr = factory.createSectPr();
        body.setSectPr(sectPr);
        sectPr.setPgMar(pgMar);

        PgMar pg = factory.createSectPrPgMar();
        pg.setTop(BigInteger.valueOf(1));

        setGlobalFontFamily(GLOBAL_FONT);

        Tbl header = createHeaderTable();
        mainDoc().addObject(header);
        createBodyContent();

        Tbl footer = createFooterTable();
        setTableWidth(footer, 8700);
        mainDoc().addObject(footer);

        File docx = new File("test_doc.docx");
        wordMLPackage.save(docx);
        System.out.println("Complete word");
        WordprocessingMLPackage wordMl = WordprocessingMLPackage.load(docx);
        //        PdfSettings pdfSettings = new PdfSettings();
        convertPdf(wordMl);
        System.out.println("Complete pdf");
    }

    static void convertPdf(final WordprocessingMLPackage wordMl) throws Exception
    {
        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setWmlPackage(wordMl);

        //        String regex = ".*(arial|times).*";
        //        PhysicalFonts.setRegex(regex);
        PhysicalFont font = PhysicalFonts.get("Arial");
        Mapper fontMapper = new IdentityPlusMapper();
        wordMl.setFontMapper(fontMapper);
        if (font != null)
        {
            fontMapper.put("Times New Roman", font);
        }

        OutputStream os = new FileOutputStream(new File("test_pdf.pdf"));
        Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

        //      PdfConversion conversion = new Conversion(wordMl);
        //        conversion.output(os, pdfSettings);
    }

    private static Tbl createFooterTable()
    {
        Tbl table = factory.createTbl();

        Tr row1 = factory.createTr();
        Tc cell1 = factory.createTc();
        P para1 = factory.createP();
        R run1 = createStyledR("Nơi nhận:", true, true, true);
        para1.getContent().add(run1);
        para1.getContent().add(factory.createBr());
        para1.getContent().add(createSimpleR("- Như đề gửi;"));
        para1.getContent().add(factory.createBr());
        para1.getContent().add(createSimpleR("- Lưu TCHC, KHDN."));
        cell1.getContent().add(para1);
        row1.getContent().add(cell1);

        addStyledTableCell(row1, "GIÁM ĐỐC", true, null, JcEnumeration.CENTER);

        table.getContent().add(row1);

        return table;
    }

    private static void createBodyContent()
    {
        //first paragraph
        P p = factory.createP();
        p.getContent().add(createStyledR("Kính gửi: Công ty", true, false, false));
        setHorizontalAlignment(p, JcEnumeration.CENTER);
        mainDoc().getContent().add(p);

        //second paragraph
        String t =
            "Ngân hàng TMCP Công thương Việt Nam – Chi nhánh ...... trân trọng thông báo khoản nợ gốc và lãi đến hạn của Quý Công ty .... như sau:";
        mainDoc().addParagraphOfText(t);

        P p2 = factory.createP();
        p2.getContent().add(createStyledR("ĐTV: ", false, false, false));
        setHorizontalAlignment(p2, JcEnumeration.RIGHT);
        mainDoc().getContent().add(p2);

        createContentTable();

        String t1 =
            "Ngân hàng TMCP Công thương Việt Nam – Chi nhánh ....... ...(tự động điền) trân trọng thông báo tới Quý Công ty,  để quý Công ty được biết và chủ động trả nợ Ngân hàng đúng hạn";
        mainDoc().addParagraphOfText(t1);

        String t2 =
            "Mọi vấn đề vướng mắc Quý Công ty vui lòng liên hệ với Ngân hàng TMCP Công thương Việt Nam Chi nhánh......(tự động điền) theo địa chỉ:";
        mainDoc().addParagraphOfText(t2);

        Ind idx = new Ind();
        idx.setLeft(new BigInteger("6"));

        String t3 = "- Phòng....... (tự động điền) Điện thoại:.............. Fax:.......... (tự động điền nếu có)";
        P pb1 = factory.createP();
        pb1.getContent().add(createStyledRWithTab(t3, false, false));
        mainDoc().getContent().add(pb1);

        String t4 =
            "- Cán bộ quan hệ khách hàng:...(tự động điền)Điện thoại:...(tự động điền)           Email: .... (tự động điền))";
        P pb2 = factory.createP();
        pb2.getContent().add(createStyledRWithTab(t4, false, false));
        mainDoc().getContent().add(pb2);

    }

    private static void createContentTable()
    {
        //create table header
        Tr header = factory.createTr();
        String[] headers =
            new String[] {
                "STT",
                "Ngày hợp đồng",
                "Ngày giải ngân",
                "STK",
                "Dư nợ hiện tại",
                "Ngày đến hạn gốc",
                "Ngày đến hạn lãi",
                "Gốc đến hạn",
                "Lãi đến hạn",
                "Tổng"
            };
        String[] fkRow1 = new String[] {
            "1", "11-11-14", "11-11-14", "123456", "45000", "11-11-14", "11-11-14", "4550000", "770000", "12323"
        };
        for (String h : headers)
        {
            addStyledTableCell(header, h, true, FONT_SIZE, JcEnumeration.CENTER);
        }

        Tr fk = factory.createTr();
        for (String c : fkRow1)
        {
            addStyledTableCell(fk, c, false, FONT_SIZE, JcEnumeration.LEFT);
        }

        //summary row
        Tr summRow = factory.createTr();

        Tc rowLabel = factory.createTc();
        P para = factory.createP();
        para.getContent().add(createStyledR("Tổng cộng", true, false, false));
        rowLabel.getContent().add(para);
        setCellHMerge(rowLabel, 7);

        summRow.getContent().add(rowLabel);
        addStyledTableCell(summRow, "12344556", false, FONT_SIZE, JcEnumeration.CENTER);
        addStyledTableCell(summRow, "4344655", false, FONT_SIZE, JcEnumeration.CENTER);
        addStyledTableCell(summRow, "56500000", false, FONT_SIZE, JcEnumeration.CENTER);

        Tbl table = factory.createTbl();
        table.getContent().add(header);
        table.getContent().add(fk);
        table.getContent().add(summRow);
        addBorders(table);

        //        TblPr prop = table.getTblPr();
        //        CTTblPPr p = new CTTblPPr();
        //        p.setTblpXSpec(STXAlign.CENTER);
        //        prop.setTblpPr(p);
        Jc jc = table.getTblPr().getJc();
        //        System.out.println(jc.getVal());
        mainDoc().getContent().add(table);

    }

    public static Tbl createHeaderTable()
    {

        Tbl table = factory.createTbl();
        List<Object> ctn = table.getContent();

        Tr row1 = factory.createTr();

        addStyledTableCell(row1, "NG\u00c2N H\u00c0NG TMCP C\u00d4NG TH\u01af\u01a0NG VI\u1ec6T NAM", false, null,
            JcEnumeration.CENTER);
        addStyledTableCell(row1, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, null, JcEnumeration.CENTER);

        Tr row2 = factory.createTr();
        addStyledTableCell(row2, "CHI NHÁNH", true, null, JcEnumeration.CENTER);
        addStyledTableCell(row2, "Độc lập- Tự do- Hạnh phúc", true, null, JcEnumeration.CENTER);

        Tr row3 = factory.createTr();
        addStyledTableCell(row3, "Số", false, null, JcEnumeration.CENTER);
        addTableCell(row3, "");

        Calendar c = Calendar.getInstance();
        Tr row4 = factory.createTr();
        addStyledTableCell(row4, "V/v Thông báo nợ đến hạn", false, null, JcEnumeration.CENTER);
        addTableCell(row4, "Hà Nội, ngày...tháng....năm " + c.get(Calendar.YEAR));

        ctn.add(row1);
        ctn.add(row2);
        ctn.add(row3);
        ctn.add(row4);

        return table;
    }

    static R createSimpleR(final String text)
    {
        return createStyledR(text, false, false, false);
    }

    static R createStyledR(final String text, final boolean bold, final boolean italic, final boolean underline)
    {
        RPr rpr = factory.createRPr();
        if (bold)
            addBoldStyle(rpr);
        if (italic)
            addItalicStyle(rpr);
        if (underline)
            addUnderlineStyle(rpr);
        Text t = factory.createText();
        t.setValue(text);
        R r = factory.createR();
        r.getContent().add(t);
        r.setRPr(rpr);
        return r;
    }

    static R createStyledRWithTab(final String text, final boolean bold, final boolean italic)
    {
        RPr rpr = factory.createRPr();
        if (bold)
            addBoldStyle(rpr);
        if (italic)
            addItalicStyle(rpr);
        Text t = factory.createText();
        t.setValue(text);
        R r = factory.createR();
        r.getContent().add(factory.createRTab());
        r.getContent().add(t);
        r.setRPr(rpr);
        return r;
    }

    static void setGlobalFontFamily(final String fontName)
    {
        @SuppressWarnings("deprecation")
        Styles styles = mainDoc().getStyleDefinitionsPart().getJaxbElement();
        for (Style s : styles.getStyle())
        {
            //            System.out.println(s.getName().getVal());
            //            if (s.getName().getVal().equalsIgnoreCase("Normal"))
            //            {
            RPr rpr = s.getRPr();
            if (rpr == null)
            {
                rpr = factory.createRPr();
                s.setRPr(rpr);
            }
            RFonts rf = rpr.getRFonts();
            if (rf == null)
            {
                rf = factory.createRFonts();
                rpr.setRFonts(rf);
            }
            rf.setAscii(fontName);
            rf.setHAnsi(fontName);
            //            }
        }
    }

    static void setCellHMerge(final Tc tableCell, final int horizontalMergedCells)
    {
        if (horizontalMergedCells > 1)
        {
            TcPr tableCellProperties = tableCell.getTcPr();
            if (tableCellProperties == null)
            {
                tableCellProperties = new TcPr();
                tableCell.setTcPr(tableCellProperties);
            }

            GridSpan gridSpan = new GridSpan();
            gridSpan.setVal(new BigInteger(String.valueOf(horizontalMergedCells)));

            tableCellProperties.setGridSpan(gridSpan);
            tableCell.setTcPr(tableCellProperties);
        }
    }

    static MainDocumentPart mainDoc()
    {
        return wordMLPackage.getMainDocumentPart();
    }

    private static void addTableCell(final Tr tableRow, final String content)
    {
        Tc tableCell = factory.createTc();
        tableCell.getContent().add(mainDoc().createParagraphOfText(content));
        tableRow.getContent().add(tableCell);
    }

    /**
     * In this method we create a table cell,add the styling and add the cell to the table row.
     */
    private static void addStyledTableCell(final Tr tableRow, final String content, final boolean bold,
        final String fontSize, final JcEnumeration alignment)
    {
        Tc tableCell = factory.createTc();
        addStyling(tableCell, content, bold, fontSize, alignment);
        tableRow.getContent().add(tableCell);
    }

    /**
     * This is where we add the actual styling information. In order to do this we first create a paragraph. Then we
     * create a text with the content of the cell as the value. Thirdly, we create a so-called run, which is a container
     * for one or more pieces of text having the same set of properties, and add the text to it. We then add the run to
     * the content of the paragraph. So far what we've done still doesn't add any styling. To accomplish that, we'll
     * create run properties and add the styling to it. These run properties are then added to the run. Finally the
     * paragraph is added to the content of the table cell.
     */
    private static void addStyling(final Tc tableCell, final String content, final boolean bold, final String fontSize,
        final JcEnumeration alignment)
    {
        P paragraph = factory.createP();

        Text text = factory.createText();
        text.setValue(content);

        R run = factory.createR();
        run.getContent().add(text);

        paragraph.getContent().add(run);

        RPr runProperties = factory.createRPr();
        if (bold)
        {
            addBoldStyle(runProperties);
        }

        if (fontSize != null && !fontSize.isEmpty())
        {
            setFontSize(runProperties, fontSize);
        }

        if (alignment != null)
            setHorizontalAlignment(paragraph, alignment);

        run.setRPr(runProperties);

        tableCell.getContent().add(paragraph);
    }

    /**
     * In this method we're going to add the font size information to the run properties. First we'll create a
     * half-point measurement. Then we'll set the fontSize as the value of this measurement. Finally we'll set the
     * non-complex and complex script font sizes, sz and szCs respectively.
     */
    private static void setFontSize(final RPr runProperties, final String fontSize)
    {
        HpsMeasure size = new HpsMeasure();
        size.setVal(new BigInteger(fontSize));
        runProperties.setSz(size);
        runProperties.setSzCs(size);
    }

    private void setCellWidth(final Tc tableCell, final int width)
    {
        if (width > 0)
        {
            TcPr tableCellProperties = tableCell.getTcPr();
            if (tableCellProperties == null)
            {
                tableCellProperties = new TcPr();
                tableCell.setTcPr(tableCellProperties);
            }
            TblWidth tableWidth = new TblWidth();
            tableWidth.setType("dxa");
            tableWidth.setW(BigInteger.valueOf(width));
            tableCellProperties.setTcW(tableWidth);
        }
    }

    private static Tbl setTableWidth(final Tbl tbl, final long width)
    {
        TblPr tblPr = factory.createTblPr();
        TblWidth tblWidth = factory.createTblWidth();
        tblWidth.setW(BigInteger.valueOf(width));
        tblWidth.setType("dxa"); // twips
        tblPr.setTblW(tblWidth);
        tbl.setTblPr(tblPr);
        return tbl;
    }

    /**
     * In this method we'll add the bold property to the run properties. BooleanDefaultTrue is the Docx4j object for the
     * b property. Technically we wouldn't have to set the value to true, as this is the default.
     */
    private static void addBoldStyle(final RPr runProperties)
    {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        runProperties.setB(b);
    }

    private static void addItalicStyle(final RPr runProperties)
    {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        runProperties.setI(b);
    }

    private static void addUnderlineStyle(final RPr runProperties)
    {
        U val = new U();
        val.setVal(UnderlineEnumeration.SINGLE);
        runProperties.setU(val);
    }

    /**
     * In this method, we once again create a regular cell, as in the previous examples.
     */
    private static void addRegularTableCell(final Tr tableRow, final String content)
    {
        Tc tableCell = factory.createTc();
        tableCell.getContent().add(wordMLPackage.getMainDocumentPart().createParagraphOfText(content));
        tableRow.getContent().add(tableCell);
    }

    /**
     * In this method we'll add the borders to the table.
     */
    private static void addBorders(final Tbl table)
    {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }

    private static void setHorizontalAlignment(final P paragraph, final JcEnumeration hAlign)
    {
        if (hAlign != null)
        {
            PPr pprop = new PPr();
            Jc align = new Jc();
            align.setVal(hAlign);
            pprop.setJc(align);
            paragraph.setPPr(pprop);
        }
    }

    class Model
    {
        private String customer;
        private String department;
        private String unit;
    }
}
