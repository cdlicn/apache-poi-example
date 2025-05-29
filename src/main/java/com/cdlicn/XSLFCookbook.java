package com.cdlicn;

import org.apache.commons.io.IOUtils;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author cdlicn
 * @date 2025年05月29日 9:55
 * @description
 */

public class XSLFCookbook {
    public static void main(String[] args) throws Exception {
        newPresentation();
        readAnExistingPresentationAndAppendASlideToIt();
        createANewSlideFromAPredefinedSlideLayout();
        deleteSlide();
        reOrderSlides();
        retrieveOrChangeSlideSize();
        readShapesContainedInAParticularSlide();
        addImageToSlide();
        readImagesContainedWithinAPresentation();
        basicTextFormatting();
        createAHyperlink();
        mergeMultiplePresentationsTogether();
    }

    public static void newPresentation() throws IOException {
        // 新建一个空的ppt
        XMLSlideShow ppt = new XMLSlideShow();
        // 添加一页幻灯片
        XSLFSlide blankSlide = ppt.createSlide();

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow1.pptx");
        ppt.write(out);
        out.close();
    }

    public static void readAnExistingPresentationAndAppendASlideToIt() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow1.pptx"));

        // 在末尾新增一页
        XSLFSlide blankSlide = ppt.createSlide();

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow2.pptx");
        ppt.write(out);
        out.close();
    }

    public static void createANewSlideFromAPredefinedSlideLayout() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow3_test.pptx"));

        // 查看可用的幻灯片布局
        System.out.println("Available slide layouts:");
        for (XSLFSlideMaster master : ppt.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                System.out.println(layout.getType());
            }
        }

        // 空白幻灯片
        XSLFSlide blankSlide = ppt.createSlide();

        // 可以有多个master，每个master引用多个layout
        // 出于演示的目的，我们使用第一个（默认）幻灯片母版
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

        // 仅标题幻灯片
        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);
        // 填充占位符
        XSLFSlide slide1 = ppt.createSlide(titleLayout);
        XSLFTextShape title1 = slide1.getPlaceholder(0);
        title1.setText("First Title");

        // 标题和内容
        XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        XSLFSlide slide2 = ppt.createSlide(titleBodyLayout);
        XSLFTextShape title2 = slide2.getPlaceholder(0);
        title2.setText("Second Title");
        XSLFTextShape body2 = slide2.getPlaceholder(1);
        // 清除所有已存在的文本
        body2.clearText();
        body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
        body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
        body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow3.pptx");
        ppt.write(out);
        out.close();
    }

    public static void deleteSlide() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow_test.pptx"));
        // 删除第一页幻灯片
        ppt.removeSlide(0);

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow4.pptx");
        ppt.write(out);
        out.close();
    }

    public static void reOrderSlides() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow_test.pptx"));
        List<XSLFSlide> slides = ppt.getSlides();

        XSLFSlide thirdSlide = slides.get(2);
        // 将第三页幻灯片移到第一页
        ppt.setSlideOrder(thirdSlide, 0);

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow5.pptx");
        ppt.write(out);
        out.close();
    }

    public static void retrieveOrChangeSlideSize() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        // 检索页面大小
        Dimension pgsize = ppt.getPageSize();
        int pgx = pgsize.width;
        int pgy = pgsize.height;
        System.out.println("Page size: " + pgx + " x " + pgy);

        // 设置新页面的大小
        ppt.setPageSize(new Dimension(1024, 768));

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow6.pptx");
        ppt.write(out);
        out.close();
    }

    public static void readShapesContainedInAParticularSlide() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow_test.pptx"));
        // 获取幻灯片
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape sh : slide.getShapes()) {
                // shape的名字
                String name = sh.getShapeName();
                System.out.println("shape name：" + name);
                // shape在幻灯片中的位置
                if (sh instanceof PlaceableShape) {
                    Rectangle2D anchor = ((PlaceableShape) sh).getAnchor();
                    System.out.println("anchor:" + anchor);
                }
                if (sh instanceof XSLFConnectorShape) {
                    XSLFConnectorShape line = (XSLFConnectorShape) sh;
                    // 使用Line.....
                    System.out.println("line：" + line);
                } else if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;
                    // 使用可以容纳文本的shape......
                    System.out.println("text：" + shape);
                } else if (sh instanceof XSLFPictureShape) {
                    XSLFPictureShape shape = (XSLFPictureShape) sh;
                    // 使用Picture......
                    System.out.println("picture：" + shape);
                }
            }
        }
    }

    public static void addImageToSlide() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        byte[] pictureData = IOUtils.toByteArray(new FileInputStream("output/XSLF/images/pic3.png"));

        XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
        XSLFPictureShape pic = slide.createPicture(pd);

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow7.pptx");
        ppt.write(out);
        out.close();
    }

    public static void readImagesContainedWithinAPresentation() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("output/XSLF/slideshow8_test.pptx"));
        for (XSLFPictureData data : ppt.getPictureData()) {
            byte[] bytes = data.getData();
            String fileName = data.getFileName();
            FileOutputStream out = new FileOutputStream("output/XSLF/down_images/" + fileName);
            out.write(bytes);
            out.close();
        }
    }

    public static void basicTextFormatting() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        XSLFTextBox shape = slide.createTextBox();
        shape.setAnchor(new Rectangle(100, 100, 500, 300));
        XSLFTextParagraph p = shape.addNewTextParagraph();

        XSLFTextRun r1 = p.addNewTextRun();
        r1.setText("The");
        r1.setFontColor(Color.BLUE);
        r1.setFontSize(24.);

        XSLFTextRun r2 = p.addNewTextRun();
        r2.setText(" quick");
        r2.setFontColor(Color.RED);
        r2.setBold(true);

        XSLFTextRun r3 = p.addNewTextRun();
        r3.setText(" brown");
        r3.setFontSize(12.);
        r3.setItalic(true);
        r3.setStrikethrough(true);

        XSLFTextRun r4 = p.addNewTextRun();
        r4.setText(" fox");
        r4.setUnderlined(true);

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow9.pptx");
        ppt.write(out);
        out.close();
    }

    public static void createAHyperlink() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        // 将超链接分配给文本
        XSLFTextBox shape = slide.createTextBox();
        shape.setAnchor(new Rectangle(100, 100, 500, 300));
        XSLFTextRun r = shape.addNewTextParagraph().addNewTextRun();
        r.setText("Apache POI");
        XSLFHyperlink link = r.createHyperlink();
        link.setAddress("https://poi.apache.org");

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow10.pptx");
        ppt.write(out);
        out.close();
    }

    public static void mergeMultiplePresentationsTogether() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        String[] inputs = {"output/XSLF/slideshow11_test1.pptx","output/XSLF/slideshow11_test2.pptx"};
        for (String arg : inputs) {
            FileInputStream is = new FileInputStream(arg);
            XMLSlideShow src = new XMLSlideShow(is);
            is.close();

            for (XSLFSlide srcSlide : src.getSlides()) {
                ppt.createSlide().importContent(srcSlide);
            }
        }

        FileOutputStream out = new FileOutputStream("output/XSLF/slideshow11.pptx");
        ppt.write(out);
        out.close();
    }

}
