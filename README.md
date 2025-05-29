# HSLF Cookbook

## 创建新演示文稿并向其中添加新幻灯片

```java
public static void newPresentation() throws IOException {
    // 创建新的空幻灯片
    HSLFSlideShow ppt = new HSLFSlideShow();

    // 添加第一张幻灯片
    HSLFSlide s1 = ppt.createSlide();

    // 添加第二章幻灯片
    HSLFSlide s2 = ppt.createSlide();

    // 在文件中保存更改
    FileOutputStream out = new FileOutputStream("output/slidershow.ppt");
    ppt.write(out);
    out.close();
}
```



## 检索或更改幻灯片大小

```java
public static void retrieveOrChangeSlideSize() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/slidershow.ppt"));

    // 获取页面大小，坐标以点表示
    Dimension pgSize = ppt.getPageSize();
    int pgx = pgSize.width;
    int pgy = pgSize.height;
    System.out.println("pgx: " + pgx); // 720
    System.out.println("pgy: " + pgy); // 540

    // 设置新的页面大小
    ppt.setPageSize(new Dimension(1024, 768));

    // 保存更改
    FileOutputStream out = new FileOutputStream("output/slidershow2.ppt");
    ppt.write(out);
    out.close();
}
```



## 在幻灯片上绘制形状

添加形状时，通常指定形状的尺寸以及形状边框左上角相对于幻灯片左上角的位置。图形图层中的距离以point为单位测量 (72 points = 1 inch).。

执行` setBold、setItalic、setUnderlined`日志出现`ERROR HSLFTextRun Sheet is not available`的错误，可以忽略

```java
public static void drawingAShapeOnASlide() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow();

    HSLFSlide s1 = ppt.createSlide();
    HSLFSlide s2 = ppt.createSlide();
    HSLFSlide s3 = ppt.createSlide();
    HSLFSlide s4 = ppt.createSlide();

    // Line shape
    HSLFLine line = new HSLFLine();
    // 设置锚点
    line.setAnchor(new Rectangle(50,50,100,20));
    // 设置颜色
    line.setLineColor(new Color(0,128,0));
    // 设置线宽
    line.setLineCompound(StrokeStyle.LineCompound.DOUBLE);
    s1.addShape(line);

    // TextBox
    HSLFTextBox txt = new HSLFTextBox();
    // 文本内容
    txt.setText("Hello,World");
    // 设置锚点
    txt.setAnchor(new Rectangle(300,100,300,500));
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
    autoShape1.setAnchor(new Rectangle(50,50,100,200));
    // 填充颜色
    autoShape1.setFillColor(Color.RED);
    s3.addShape(autoShape1);

    // Trapezoid
    HSLFAutoShape autoShape2 = new HSLFAutoShape(ShapeType.TRAPEZOID);
    autoShape2.setAnchor(new Rectangle(150,150,100,200));
    autoShape2.setFillColor(Color.BLUE);
    s4.addShape(autoShape2);

    FileOutputStream out = new FileOutputStream("output/slideshow3.ppt");
    ppt.write(out);
    out.close();
}
```



## 获取特定幻灯片中包含的形状

```java
public static void shapesContainedInAParticularSlide() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/slideshow3.ppt"));
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
```



## 如何使用图片

目前，HSLF API支持以下类型的图片：

- Windows图元文件（WMF）
- 增强的图元文件（EMF）
- JPEG交换格式
- 可移植网络图形（PNG）
- Macintosh PICT

【测试后发现，ppt插入png格式图片会出现图片破损情况，换成jpg格式即可】

```java
public static void workWithPictures() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/slideshow4_pictest.ppt"));

    // 将新图片插入到新幻灯片中
    HSLFPictureData pd = ppt.addPicture(new File("output/images/pic1.jpg"), PictureData.PictureType.JPEG);
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
        FileOutputStream out = new FileOutputStream("output/down_images/pict_" + idx + ext);
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
            FileOutputStream out = new FileOutputStream("output/down_images/sslide0_" + idx + type.extension);
            out.write(data);
            out.close();
            idx++;
        }
    }

    FileOutputStream out = new FileOutputStream("output/slideshow4.ppt");
    ppt.write(out);
    out.close();
}
```



## 设置幻灯片标题

```java
public static void setSlideTitle() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow();
    HSLFSlide slide = ppt.createSlide();
    HSLFTextBox title = slide.addTitle();
    title.setText("Hello,World");

    // 保存修改
    FileOutputStream out = new FileOutputStream("output/slideshow5.ppt");
    ppt.write(out);
    out.close();
}
```



## 修改幻灯片母版的背景

`HSLFFill.FILL_PICTURE`为私有，不可直接访问

```java
public static void modifyBackgroundOfASlideMaster() throws IOException, ClassNotFoundException {
    HSLFSlideShow ppt = new HSLFSlideShow();
    HSLFSlideMaster master = ppt.getSlideMasters().get(0);

    HSLFFill fill = master.getBackground().getFill();
    HSLFPictureData pd = ppt.addPicture(new File("output/images/pic1.jpg"), PictureData.PictureType.JPEG);
    //        fill.setFillType(HSLFFill.FILL_PICTURE);
    fill.setFillType(3);
    fill.setPictureData(pd);

    // 保存修改
    FileOutputStream out = new FileOutputStream("output/slideshow6.ppt");
    ppt.write(out);
    out.close();
}
```



## 修改幻灯片的背景

```java
public static void modifyBackgroundOfASlide() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow();
    HSLFSlide slide = ppt.createSlide();

    // 假设这个幻灯片有自己的背景，如果没有这一行，将使用master的背景
    slide.setFollowMasterBackground(false);
    HSLFFill fill = slide.getBackground().getFill();
    HSLFPictureData pd = ppt.addPicture(new File("output/images/pic1.jpg"), PictureData.PictureType.JPEG);
    //        fill.setFillType(HSLFFill.FILL_PICTURE);
    fill.setFillType(3);
    fill.setPictureData(pd);

    FileOutputStream out = new FileOutputStream("output/slideshow7.ppt");
    ppt.write(out);
    out.close();
}
```



## 修改形状的背景

```java
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

    FileOutputStream out = new FileOutputStream("output/slideshow8.ppt");
    ppt.write(out);
    out.close();
}
```



## 创建列表

```java
public static void createBulletedLists() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow();

    HSLFSlide slide = ppt.createSlide();

    HSLFTextBox shape = new HSLFTextBox();
    HSLFTextParagraph tp = shape.getTextParagraphs().get(0);
    // 设置Bullet
    tp.setBullet(true);
    // 设置列表前的符号
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

    shape.setAnchor(new Rectangle(100,100,500,300));
    slide.addShape(shape);

    FileOutputStream out = new FileOutputStream("output/slideshow9.ppt");
    ppt.write(out);
    out.close();
}
```



## 从幻灯片中读取超链接

```java
public static void readHyperlinksFromASlideShow() throws IOException {
    FileInputStream is = new FileInputStream("output/slideshow10_linktest.ppt");
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
```



## 创建表

`此版本不存在 table.createBorder()方法`

`该方法会在同一个位置创建两个表格，slide.createTable(5, 2)会在创建表格的同时将表格加入幻灯片，slide.addShape(table)会将表格重复加入幻灯片。`

`如果最终不使用slide.addShape(table)方法，会导致表格边框样式缺失`

```java
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

    slide.addShape(table);
    table.moveTo(100,100);

    FileOutputStream out = new FileOutputStream("output/slideshow11.ppt");
    ppt.write(out);
    out.close();
}
```



## 从幻灯片中删除形状

```java
public static void removeShapesFromASlide() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("output/slideshow12_shapetest.ppt"));
    List<HSLFSlide> slides = ppt.getSlides();
    for (HSLFSlide slide : slides) {
        for (HSLFShape shape : slide.getShapes()) {
            // 移除这个形状
            boolean ok = slide.removeShape(shape);
            if(ok) {
                System.out.println("shape removed");
            }
        }
    }

    FileOutputStream out = new FileOutputStream("output/slideshow12.ppt");
    ppt.write(out);
    out.close();
}
```



## 检索嵌入的OLE对象

OLE：**将其他应用程序中的对象嵌入到幻灯片中**

`当前版本HSLFShape 子类未找到OLEShape`。如需要该方法，尝试老版本



## 检索嵌入的声音

`.ppt`插入MP3格式音频时，会要求保存为`.pptx`格式，如果不想保存，可以考虑使用`.wav`格式的音频

```java
public static void retrieveEmbeddedSounds() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow(new FileInputStream("output/slideshow14_soundtest.ppt"));
    for (HSLFSoundData sound : ppt.getSoundData()) {
        // 在磁盘上保存WAV音频
        if (sound.getSoundType().equals(".WAV")){
            FileOutputStream out = new FileOutputStream("output/" + sound.getSoundName() + "_down.wav");
            out.write(sound.getData());
            out.close();
        }
    }
}
```



## 创建任意几何形状

```java
public static void createShapesOfArbitraryGeometry() throws IOException {
    HSLFSlideShow ppt = new HSLFSlideShow();
    HSLFSlide slide = ppt.createSlide();

    GeneralPath path = new GeneralPath();
    path.moveTo(100, 100);
    path.lineTo(200, 100);
    // 曲线
    path.curveTo(50, 45, 134, 22, 78, 133);
    path.lineTo(100,200);
    path.closePath();

    HSLFFreeformShape shape = new HSLFFreeformShape();
    shape.setPath(path);
    slide.addShape(shape);

    FileOutputStream out = new FileOutputStream("output/slideshow15.ppt");
    ppt.write(out);
    out.close();
}
```



## 使用Graphics2D绘制幻灯片

<span style="color:red">警告：</span>PowerPoint Graphics2D驱动程序的当前实现不完全符合java.awt.Graphics2D规范。一些功能，如剪切，绘图的图像还不支持。

`当前版本缺少new PPGraphics2D(group)对象`。如果需要可以尝试切换版本



## 将PowerPoint幻灯片导出到java.awt.Graphics2D

HSLF提供了一种将幻灯片导出为图像的方法。您可以将幻灯片捕获到java.awt.Graphics2D对象（或任何其他对象）中，并将其序列化为PNG或JPEG格式。请注意，尽管HSLF尝试尽可能接近PowerPoint渲染幻灯片，但由于以下原因，输出可能与PowerPoint不同：

- Java2D呈现字体与PowerPoint不同。字体字形的绘制方式总是有些不同
- HSLF使用java.awt.font.LineBreakMeasurer将文本拆分为多行。PowerPoint可能会以不同的方式来做。
- 如果演示文稿中的字体不可用，则将使用JDK默认字体。

当前限制：

- 尚不支持某些类型的形状（艺术字、复杂的自动形状）
- 只有位图图像（PNG、JPEG、DIB）可以在Java中呈现

```java
public static void ExportPowerPointSlidesIntoGraphics2D() throws IOException {
    FileInputStream is = new FileInputStream("output/slideshow17_test.ppt");
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
        FileOutputStream out = new FileOutputStream("output/slide-" + idx + ".png");
        ImageIO.write(img, "png", out);
        out.close();
        idx++;
    }
}
```



## 如何设置页眉/页脚

```java
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
```



## 从现有演示文稿中提取页眉/页脚

```java
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
```

































































































































