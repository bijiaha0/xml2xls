
import jxl.write.Label;
import jxl.write.WritableSheet;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import xml.Xml;

import java.io.File;
import java.net.MalformedURLException;

/**
 * author： nero
 * email: nerosoft@outlook.com
 * data: 16-9-20
 * time: 下午11:58.
 */
public class Test {

    public static void main(String[] args) {
        File xlsFile = new File("/Users/BJH/Desktop/xml2xls-finally/xml2xls-master/src/test/resources/test.xls");
        Xml xml = new Xml();//xml资源加载器管理器
        Xml2Xls xml2Xls = new Xml2Xls(xlsFile, "sheet"); //转换器
        xml2Xls.init();
        int row = 0;

        File xmlFileHead = new File("/Users/BJH/Desktop/xml2xls-finally/xml2xls-master/src/test/resources/cancer (1).xml");
        try {
            for (Element element : xml.getAllElement(xml.getRootElement(xml.read(xmlFileHead)))) {
                int column = 0;
                //System.out.println(element.getName());
                if (element.getName() == "admin") {
                    //row-=1;
                    continue;
                }
                for (Element childElement : xml.getAllElement(element)) {
                    //System.out.print(" |" + childElement.getText());
                    System.out.println(childElement.getName());

                    xml2Xls.in2xls(new Label(column, row, childElement.getName()));
                    column++;

                }
                //System.out.println(" |");
                row++;
            }
        }catch(Exception e){
            e.printStackTrace();

        }

        for (int i = 1; i < 1097; i++) {
            File xmlFile = new File("/Users/BJH/Desktop/xml2xls-finally/xml2xls-master/src/test/resources/cancer ("+i+").xml");

            try {
                for (Element element : xml.getAllElement(xml.getRootElement(xml.read(xmlFile)))) {
                    int column = 0;
                    //System.out.println(element.getName());
                    if (element.getName() == "admin") {
                        //row-=1;
                        continue;
                    }
                    for (Element childElement : xml.getAllElement(element)) {
                        //System.out.print(" |" + childElement.getText());
                        //System.out.println(childElement.getName());

                        xml2Xls.in2xls(new Label(column, row, childElement.getText()));
                        column++;

                    }
                    //System.out.println(" |");
                    row++;
                }
            } catch (MalformedURLException e) {
                e.printStackTrace();
            } catch (DocumentException e) {
                e.printStackTrace();
            }
        }
        xml2Xls.exit();
    }

}
