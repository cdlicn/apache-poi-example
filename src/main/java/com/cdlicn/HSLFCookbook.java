package com.cdlicn;

import com.sun.net.httpserver.Headers;
import org.apache.poi.hslf.model.HeadersFooters;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.sl.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.GeneralPath;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;

/**
 * @author cdlicn
 * @date 2025年05月27日 15:13
 * @description
 */

public class HSLFCookbook {


    public static void main(String[] args) throws Exception {
        newPresentation();
        retrieveOrChangeSlideSize();
        drawingAShapeOnASlide();
        shapesContainedInAParticularSlide();
        workWithPictures();
        setSlideTitle();
        modifyBackgroundOfASlideMaster();
        modifyBackgroundOfASlide();
        modifyBackgroundOfAShape();
        createBulletedLists();
        readHyperlinksFromASlideShow();
        createTables();
        removeShapesFromASlide();

        retrieveEmbeddedSounds();
        createShapesOfArbitraryGeometry();

        ExportPowerPointSlidesIntoGraphics2D();
        setHeadersOrFooters();
        extractHeadersOrFootersFromAnExistingPresentation();
    }


    public static void newPresentation() throws IOException {
        // 创建新的空幻灯片
        HSLFSlideShow ppt = new HSLFSlideShow();

        // 添加第一张幻灯片
        HSLFSlide s1 = ppt.createSlide();

        // 添加第二章幻灯片
        HSLFSlide s2 = ppt.createSlide();

        // 在文件中保存更改
        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow.ppt");
        ppt.write(out);
        out.close();
    }

    public static void retrieveOrChangeSlideSize() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/HSLF/slideshow.ppt"));

        // 获取页面大小，坐标以点表示
        Dimension pgSize = ppt.getPageSize();
        int pgx = pgSize.width;
        int pgy = pgSize.height;
        System.out.println("pgx: " + pgx); // 720
        System.out.println("pgy: " + pgy); // 540

        // 设置新的页面大小
        ppt.setPageSize(new Dimension(1024, 768));

        // 保存更改
        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow2.ppt");
        ppt.write(out);
        out.close();
    }

    public static void drawingAShapeOnASlide() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();

        HSLFSlide s1 = ppt.createSlide();
        HSLFSlide s2 = ppt.createSlide();
        HSLFSlide s3 = ppt.createSlide();
        HSLFSlide s4 = ppt.createSlide();

        // Line shape
        HSLFLine line = new HSLFLine();
        // 设置锚点
        line.setAnchor(new Rectangle(50, 50, 100, 20));
        // 设置颜色
        line.setLineColor(new Color(0, 128, 0));
        // 设置线宽
        line.setLineCompound(StrokeStyle.LineCompound.DOUBLE);
        s1.addShape(line);

        // TextBox
        HSLFTextBox txt = new HSLFTextBox();
        // 文本内容
        txt.setText("Hello,World");
        // 设置锚点
        txt.setAnchor(new Rectangle(300, 100, 300, 500));
        // 使用TextRun处理文本格式
        HSLFTextParagraph tp = txt.getTextParagraphs().get(0);
        // 设置对齐方式
        tp.setTextAlign(TextParagraph.TextAlign.RIGHT);
        // 设置字体样式
        HSLFTextRun rt = tp.getTextRuns().get(0);
        // 设置字体大小
        rt.setFontSize(36D);
        // 设置字体
        rt.setFontFamily("楷书");
        // 加粗
        rt.setBold(true);
        // 斜体
        rt.setItalic(true);
        // 下划线
        rt.setUnderlined(true);
        // 颜色
        rt.setFontColor(Color.RED);
        s2.addShape(txt);

        // AutoShape
        // 32-point star
        HSLFAutoShape autoShape1 = new HSLFAutoShape(ShapeType.STAR_32);
        // 设置锚点
        autoShape1.setAnchor(new Rectangle(50, 50, 100, 200));
        // 填充颜色
        autoShape1.setFillColor(Color.RED);
        s3.addShape(autoShape1);

        // Trapezoid
        HSLFAutoShape autoShape2 = new HSLFAutoShape(ShapeType.TRAPEZOID);
        autoShape2.setAnchor(new Rectangle(150, 150, 100, 200));
        autoShape2.setFillColor(Color.BLUE);
        s4.addShape(autoShape2);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow3.ppt");
        ppt.write(out);
        out.close();
    }

    public static void shapesContainedInAParticularSlide() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/HSLF/slideshow3.ppt"));
        // 获取幻灯片
        for (HSLFSlide slide : ppt.getSlides()) {
            for (HSLFShape sh : slide.getShapes()) {
                // 形状的名称
                String name = sh.getShapeName();
                System.out.println("shapeName: " + name);

                // 形状的锚点，用于定义此形状在幻灯片中的位置
                Rectangle2D anchor = sh.getAnchor();
                System.out.println(anchor);
                if (sh instanceof HSLFLine) {
                    HSLFLine line = (HSLFLine) sh;
                    // 使用Line...
                    System.out.println(line);
                } else if (sh instanceof HSLFAutoShape) {
                    HSLFAutoShape shape = (HSLFAutoShape) sh;
                    // 使用AutoShape...
                    System.out.println(shape);
                } else if (sh instanceof HSLFTextBox) {
                    HSLFTextBox shape = (HSLFTextBox) sh;
                    // 使用TextBox...
                    System.out.println(shape);
                } else if (sh instanceof HSLFPictureShape) {
                    HSLFPictureShape shape = (HSLFPictureShape) sh;
                    // 使用Picture...
                    System.out.println(shape);
                }
            }
        }
    }

    public static void workWithPictures() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/HSLF/slideshow4_pictest.ppt"));

        // 将新图片插入到新幻灯片中
        HSLFPictureData pd = ppt.addPicture(new File("output/HSLF/images/pic1.jpg"), PictureData.PictureType.JPEG);
        HSLFPictureShape picNew = new HSLFPictureShape(pd);
        // 在幻灯片中设置图像位置
        picNew.setAnchor(new Rectangle(100, 100, 300, 200));
        HSLFSlide slide = ppt.createSlide();
        slide.addShape(picNew);

        // 提取演示文稿中包含的所有图片，并保存
        int idx = 1;
        for (HSLFPictureData pict : ppt.getPictureData()) {
            //picture data
            byte[] data = pict.getData();

            PictureData.PictureType type = pict.getType();
            String ext = type.extension;
            FileOutputStream out = new FileOutputStream("output/HSLF/down_images/pict_" + idx + ext);
            out.write(data);
            out.close();
            idx++;
        }

        // 检索第一张幻灯片中包含的图片，并将其保存在磁盘上
        idx = 1;
        slide = ppt.getSlides().get(0);
        for (HSLFShape sh : slide.getShapes()) {
            if (sh instanceof HSLFPictureShape) {
                HSLFPictureShape pict = (HSLFPictureShape) sh;
                HSLFPictureData pictData = pict.getPictureData();
                byte[] data = pictData.getData();
                PictureData.PictureType type = pictData.getType();
                FileOutputStream out = new FileOutputStream("output/HSLF/down_images/sslide0_" + idx + type.extension);
                out.write(data);
                out.close();
                idx++;
            }
        }

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow4.ppt");
        ppt.write(out);
        out.close();
    }

    public static void setSlideTitle() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlide slide = ppt.createSlide();
        HSLFTextBox title = slide.addTitle();
        title.setText("Hello,World");

        // 保存修改
        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow5.ppt");
        ppt.write(out);
        out.close();
    }

    public static void modifyBackgroundOfASlideMaster() throws IOException, ClassNotFoundException {
        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlideMaster master = ppt.getSlideMasters().get(0);

        HSLFFill fill = master.getBackground().getFill();
        HSLFPictureData pd = ppt.addPicture(new File("output/HSLF/images/pic1.jpg"), PictureData.PictureType.JPEG);
//        fill.setFillType(HSLFFill.FILL_PICTURE);
        fill.setFillType(3);
        fill.setPictureData(pd);

        // 保存修改
        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow6.ppt");
        ppt.write(out);
        out.close();
    }

    public static void modifyBackgroundOfASlide() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlide slide = ppt.createSlide();

        // 假设这个幻灯片有自己的背景，如果没有这一行，将使用master的背景
        slide.setFollowMasterBackground(false);
        HSLFFill fill = slide.getBackground().getFill();
        HSLFPictureData pd = ppt.addPicture(new File("output/HSLF/images/pic1.jpg"), PictureData.PictureType.JPEG);
//        fill.setFillType(HSLFFill.FILL_PICTURE);
        fill.setFillType(3);
        fill.setPictureData(pd);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow7.ppt");
        ppt.write(out);
        out.close();
    }

    public static void modifyBackgroundOfAShape() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlide slide = ppt.createSlide();

        HSLFAutoShape shape = new HSLFAutoShape(ShapeType.RECT);
        shape.setAnchor(new Rectangle(100, 100, 200, 200));
        HSLFFill fill = shape.getFill();
//        fill.setFillType(HSLFFill.FILL_SHADE);
        fill.setFillType(4);
        fill.setBackgroundColor(Color.RED);
        fill.setForegroundColor(Color.GREEN);

        slide.addShape(shape);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow8.ppt");
        ppt.write(out);
        out.close();
    }

    public static void createBulletedLists() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();

        HSLFSlide slide = ppt.createSlide();

        HSLFTextBox shape = new HSLFTextBox();
        HSLFTextParagraph tp = shape.getTextParagraphs().get(0);
        // 设置Bullet
        tp.setBullet(true);
        // 项目符号
        tp.setBulletChar('\u263A');
        // 符号偏移量
        tp.setIndent(2.);
        // text偏移量（应大于项目符号偏移量）
        tp.setLeftMargin(50.);

        HSLFTextRun rt = tp.getTextRuns().get(0);
        shape.setText(
                "January\r" +
                        "February\r" +
                        "March\r" +
                        "April"
        );
        rt.setFontSize(42.);
        slide.addShape(shape);

        shape.setAnchor(new Rectangle(100, 100, 500, 300));
        slide.addShape(shape);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow9.ppt");
        ppt.write(out);
        out.close();
    }

    public static void readHyperlinksFromASlideShow() throws IOException {
        FileInputStream is = new FileInputStream("output/HSLF/slideshow10_linktest.ppt");
        HSLFSlideShow ppt = new HSLFSlideShow(is);
        is.close();

        for (HSLFSlide slide : ppt.getSlides()) {
            // 从文本串中读取超链接
            for (List<HSLFTextParagraph> txt : slide.getTextParagraphs()) {
                for (HSLFTextParagraph para : txt) {
                    for (HSLFTextRun run : para) {
                        HSLFHyperlink link = run.getHyperlink();
                        if (link != null) {
                            String title = link.getLabel();
                            String address = link.getAddress();
                            String text = run.getRawText();
                            System.out.println("title: " + title);
                            System.out.println("address: " + address);
                            System.out.println("text: " + text);
                            System.out.println("-----------------------------------------------------------------------");
                        }
                    }
                }
            }

            // 在PowerPoint中，可以将超链接分配给没有文本的shape,例如一个Line对象
            // 下面的代码演示了如何读取这样的超链接
            for (HSLFShape sh : slide.getShapes()) {
                if (sh instanceof HSLFSimpleShape) {
                    HSLFHyperlink link = ((HSLFSimpleShape) sh).getHyperlink();
                    if (link != null) {
                        String title = link.getLabel();
                        String address = link.getAddress();
                        System.out.println("title: " + title);
                        System.out.println("address: " + address);
                    }
                }
            }
        }
    }

    public static void createTables() throws IOException {
        // 表格数据
        String[][] data = {
                {"INPUT FILE", "NUMBER OF RECORDS"},
                {"Item File", "11,559"},
                {"Vendor File", "300"},
                {"Purchase History File", "10,000"},
                {"Total # of requisitions", "10,200,038"}
        };

        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlide slide = ppt.createSlide();

        // 创建一个5行2列的表格
        HSLFTable table = slide.createTable(5, 2);
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[i].length; j++) {
                HSLFTableCell cell = table.getCell(i, j);
                cell.setText(data[i][j]);
                cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
                cell.setHorizontalCentered(true);

                // 设置表格边框
                // 此版本不存在 table.createBorder()方法
                cell.setLineColor(Color.GREEN);
                cell.setLineWidth(3.);

                HSLFTextRun rt = cell.getTextParagraphs().get(0).getTextRuns().get(0);
                rt.setFontFamily("宋体");
                rt.setFontSize(10.);
            }
        }

        // 设置第一列的宽度
        table.setColumnWidth(0, 300.);
        // 设置第二列的宽度
        table.setColumnWidth(1, 150.);

        table.moveTo(100, 100);
        slide.addShape(table);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow11.ppt");
        ppt.write(out);
        out.close();
    }

    public static void removeShapesFromASlide() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/HSLF/slideshow12_shapetest.ppt"));
        List<HSLFSlide> slides = ppt.getSlides();
        for (HSLFSlide slide : slides) {
            for (HSLFShape shape : slide.getShapes()) {
                // 移除这个形状
                boolean ok = slide.removeShape(shape);
                if (ok) {
                    System.out.println("shape removed");
                }
            }
        }

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow12.ppt");
        ppt.write(out);
        out.close();
    }

    public static void retrieveEmbeddedSounds() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow(new FileInputStream("output/HSLF/slideshow14_soundtest.ppt"));
        for (HSLFSoundData sound : ppt.getSoundData()) {
            // 在磁盘上保存WAV音频
            if (sound.getSoundType().equals(".WAV")) {
                FileOutputStream out = new FileOutputStream("output/HSLF/" + sound.getSoundName() + "_down.wav");
                out.write(sound.getData());
                out.close();
            }
        }
    }

    public static void createShapesOfArbitraryGeometry() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();
        HSLFSlide slide = ppt.createSlide();

        GeneralPath path = new GeneralPath();
        path.moveTo(100, 100);
        path.lineTo(200, 100);
        // 曲线
        path.curveTo(50, 45, 134, 22, 78, 133);
        path.lineTo(100, 200);
        path.closePath();

        HSLFFreeformShape shape = new HSLFFreeformShape();
        shape.setPath(path);
        slide.addShape(shape);

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow15.ppt");
        ppt.write(out);
        out.close();
    }

    public static void ExportPowerPointSlidesIntoGraphics2D() throws IOException {
        FileInputStream is = new FileInputStream("output/HSLF/slideshow17_test.ppt");
        HSLFSlideShow ppt = new HSLFSlideShow(is);
        is.close();

        Dimension pgsize = ppt.getPageSize();
        int idx = 1;
        for (HSLFSlide slide : ppt.getSlides()) {
            BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = img.createGraphics();

            // 清除绘图区域
            graphics.setPaint(Color.white);
            graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

            // 渲染
            slide.draw(graphics);

            // 保存输出
            FileOutputStream out = new FileOutputStream("output/HSLF/slide-" + idx + ".png");
            ImageIO.write(img, "png", out);
            out.close();
            idx++;
        }
    }

    public static void setHeadersOrFooters() throws IOException {
        HSLFSlideShow ppt = new HSLFSlideShow();

        // 演示范围页眉/页脚
        HeadersFooters hdd = ppt.getSlideHeadersFooters();
        hdd.setSlideNumberVisible(true);
        hdd.setFootersText("Created by POI-HSLF");

        ppt.createSlide();
        ppt.createSlide();

        FileOutputStream out = new FileOutputStream("output/HSLF/slideshow18.ppt");
        ppt.write(out);
        out.close();
    }

    public static void extractHeadersOrFootersFromAnExistingPresentation() throws IOException {
        FileInputStream is = new FileInputStream("output/HSLF/slideshow18.ppt");
        HSLFSlideShow ppt = new HSLFSlideShow(is);
        is.close();

        // 演示范围页眉/页脚
        HeadersFooters hdd = ppt.getSlideHeadersFooters();
        if (hdd.isFooterVisible()) {
            String footerText = hdd.getFooterText();
            System.out.println("hdd-footerText：" + footerText);
        }

        // 每张幻灯片页眉/页脚
        for (HSLFSlide slide : ppt.getSlides()) {
            HeadersFooters hdd2 = slide.getHeadersFooters();
            if (hdd2.isFooterVisible()) {
                String footerText = hdd2.getFooterText();
                System.out.println("hdd2-footerText：" + footerText);
            }
            if (hdd2.isUserDateVisible()) {
                String dateTimeText = hdd2.getDateTimeText();
                System.out.println("hdd2-dateTimeText：" + dateTimeText);
            }
            if(hdd2.isSlideNumberVisible()){
                int slideNumber = slide.getSlideNumber();
                System.out.println("hdd2-slideNumberText：" + slideNumber);
            }
        }
    }

}

























