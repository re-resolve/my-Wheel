# 1、DocxToDocument

## The function of this wheel：

将word文档中的数学公式转化为latex语言、将图片转为base64编码的字符串，同时在原文中发生替换的内容中添加上左分隔符和右分隔符（可自定义）。完成替换的工作后，可在返回结果的基础上进一步添加自己的功能，如：打印输出文档每行的内容、导出word文档等等。

## Process:

```
docx file -> XWPFDocument -> Document -> Latex2Context && Picture2Context -> Document
```

源码中：

大体流程：将word文档（docx文件）读取为 XWPFDocument 对象、再进一步转为 Document 对象。然后将 Document 对象的**公式节点**和**图片节点**分别进行转化，然后再放进原节点的文本内容中。

细节流程：

1. word文档中的数学公式转化为latex

​	格式转换流程：

```
OMML -> MML -> Latex -> context(String)
```

​	2. word文档中的图片内容转化为base64编码的字符串

​	格式转换流程：

```
picture -> bytes -> Base64.toString() -> context
```

## Usage:

在idea编译器中，在 “文件” -> “项目结构” -> “模块” -> “+” -> “JAR或目录” -> 选择已下载到本地的jar包。

然后**确保**下图的红色箭头所指之处的设置没有**打勾**：

![image-20230403195437463](https://github.com/re-resolve/my-Wheel/blob/main/image-20230403195437463.png)


演示：使用jar包中的DocxToDocument类中的静态方法 docx2Document（）

该方法的介绍：

```
/**
 * 方法功能：docx file -> XWPFDocument -> Document -> Latex2Context && Picture2Context -> Document
 *
 * @param docxFile              docx文件
 * @param ommlXslPath           将OMML转化为MML的XSLT文件的路径
 * @param mmlXslPath            将MML转化为Latex的XSLT文件的路径
 * @param xslFolderPath         将MML转化为Latex的其余使用到的一些XSLT文件的文件夹路径
 * @param latexLeftSeparator    转化成latex语言后，它的左分隔符
 * @param latexRightSeparator   转化成latex语言后，它的右分隔符
 * @param pictureLeftSeparator  转化成图片的base64字符串后，它的左分隔符
 * @param pictureRightSeparator 转化成图片的base64字符串后，它的右分隔符
 * @return ommlDoc 返回的结果为内容是OMML格式的Document对象
 * @throws InvalidFormatException
 * @throws IOException
 * @throws ParserConfigurationException
 * @throws SAXException
 * @throws XPathExpressionException
 * @throws TransformerException
 */
```

将要转换的word文档：test1.docx 

如图：（Just for an example）

![image-20230403193635110](https://github.com/re-resolve/my-Wheel/blob/main/image-20230403193635110.png)

示例代码如下：

```java
    public static void main(String[] args) throws Exception {
        
        String filePath = XSLUtils.class.getResource("/test1.docx").getFile();
        
        File docxFile = new File(filePath);
    
        String ommlXslPath = "/XSLT/OMML2MML.XSL";//此处不能更改（相关使用到的xsl文件（XSLT）已存在于本jar包中）
    
        String mmlXslPath = "/XSLT/mml2tex/mmltex.xsl";//此处不能更改（相关使用到的xsl文件（XSLT）已存在于本jar包中）
    
        String xslFolderPath = "/XSLT/mml2tex/";//此处不能更改（相关使用到的xsl文件（XSLT）已存在于本jar包中）
    
        String latexLeftSeparator = "<latex>";//此处可自定义
        String latexRightSeparator = "</latex>";//此处可自定义
        String pictureLeftSeparator = "<picture>";//此处可自定义
        String pictureRightSeparator = "</picture>";//此处可自定义
        
        Document document = docx2Document(docxFile, ommlXslPath, mmlXslPath, xslFolderPath
                , latexLeftSeparator, latexRightSeparator
                , pictureLeftSeparator, pictureRightSeparator);
        //其中 local-name() 函数用于获取节点名称中的本地名称部分，避免命名空间的问题。
        XPath xpath = XPathFactory.newInstance().newXPath();
        XPathExpression expr = xpath.compile("//*[local-name()='p']");
        // 执行XPath表达式，获取所有的节点
        NodeList nodeList = (NodeList) expr.evaluate(document, XPathConstants.NODESET);
        String outputPath="C:\\Users\\Think\\Desktop\\output.docx";//打印输出的word文档路径可选
        
        //用于输出文档的对象
        XWPFDocument outputDocument = new XWPFDocument();

        for (int i = 0; i < nodeList.getLength(); i++) {
            Node node = nodeList.item(i);
            String textContent = node.getTextContent();
            System.out.println(textContent);
            outputDocument.createParagraph().createRun().setText(textContent);
        }
        FileOutputStream fos = new FileOutputStream(outputPath);
        outputDocument.write(fos);
        fos.close();
        outputDocument.close();
    }
```



## Effect demonstration:

**效果图（新生产的output.docx）：**

![image-20230403195812971](https://github.com/re-resolve/my-Wheel/blob/main/image-20230403195812971.png)

**控制台打印：**

![image-20230403195956457](https://github.com/re-resolve/my-Wheel/blob/main/image-20230403195956457.png)



## Dependencies：

建议导入：

```xml
<dependency>
   <groupId>org.apache.xmlbeans</groupId>
   <artifactId>xmlbeans</artifactId>
   <version>3.1.0</version>
</dependency>
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi</artifactId>
   <version>4.1.2</version>
</dependency>
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi-ooxml</artifactId>
   <version>4.1.2</version>
</dependency>
```

## Conclusion：

**If it helps,give me a Star!!!!!**
