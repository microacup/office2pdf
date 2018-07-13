> 本文介绍两种常用的Office文档转PDF解决方案。

## 1. [Windows平台]使用Microsoft Office

### 1.1 特点

优点：

- 兼容好；
- 速度快；

缺点：

- 只能在Windows平台使用；

### 1.2 依赖

- Microsoft Office

- [jacob](https://sourceforge.net/projects/jacob-project/)

  > JACOB is a JAVA-COM Bridge that allows you to call COM Automation components from Java. It uses JNI to make native calls to the COM libraries. JACOB runs on x86 and x64 environments supporting 32 bit and 64 bit JVMs . **需要把jacob-x.xx-x64.dll放到java/bin(与java.exe相同)目录下**

### 1.3 示例

支持office所有格式文档，每个文档转换调用接口略有区别。

**word转pdf：**

```java
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import org.junit.Test;

import java.io.File;

public class Word2PdfTest {
    
	/**
     * 测试本地Office转换word到pdf
     */
    @Test
    public void testWord() {
        String source = "d:\\tmp\\test1.doc";
        String target = "d:\\tmp\\test.pdf";

        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            File targetFile = new File(target);
            if (targetFile.exists()) {
                targetFile.delete();
            }

            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();

            System.out.println("打开文档" + source);
            doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();

            System.out.println("转换文档到PDF " + target);
            Dispatch.call(doc, "SaveAs", target, 17); // wordSaveAsPDF为特定值17

            long end = System.currentTimeMillis();
            System.out.println("转换完成用时：" + (end - start) + "ms.");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (doc != null) {
                Dispatch.call(doc, "Close", false);
            }

            if (app != null) {
                app.invoke("Quit", 0); // 不保存待定的更改
            }

            ComThread.Release();
        }
    }
}
```



**ppt 转 pdf：**

```java
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import org.junit.Test;

import java.io.File;

public class PPT2PdfTest {
    	
	/**
     * 测试本地Office转换ppt到pdf
     */
    @Test
    public void testPPT() {
        String source = "d:\\tmp\\p2.pptx";
        String target = "d:\\tmp\\wp2.pdf";

        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch ppt = null;

        try {
            File targetFile = new File(target);
            if (targetFile.exists()) {
                targetFile.delete();
            }

            ComThread.InitSTA();
            app = new ActiveXComponent("Powerpoint.Application");
            Dispatch ppts = app.getProperty("Presentations").toDispatch();

            /*
             * call
             * param 4: ReadOnly
             * param 5: Untitled指定文件是否有标题
             * param 6: WithWindow指定文件是否可见
             * */
            System.out.println("打开文档" + source);
            ppt = Dispatch.call(ppts, "Open", source, true, true, false).toDispatch();

            System.out.println("正在转换为PDF " + target);
            Dispatch.call(ppt, "SaveAs", target, 32); // pptSaveAsPDF为特定值32

            long end = System.currentTimeMillis();
            System.out.println("转换完成用时：" + (end - start) + "ms.");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }

            if (app != null) {
                app.invoke("Quit");
            }
            ComThread.Release();
        }
    }
}
```



## 2. [所有平台]使用OpenOffice

### 2.1 特点

优点：

- 支持Linux

缺点：

- 性能差；
- 兼容性不好；

### 2.2 依赖

- OpenOffice/LibreOffice 
- jodconverter

```groovy
compile (
            'org.openoffice:unoil:4.1.2',
            'org.openoffice:juh:4.1.2',
            'org.openoffice:ridl:4.1.2',
            'org.openoffice:jurt:4.1.2',
            'org.jodconverter:jodconverter-core:4.2.0',
            'org.jodconverter:jodconverter-local:4.2.0'
    )
```

### 2.3 示例

word、ppt等接口相同：

```java
import org.jodconverter.JodConverter;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;
import org.jodconverter.office.OfficeUtils;
import org.junit.Test;

import java.io.File;

/**
 * 2018/7/13
 */
public class OpenOfficeTest {
    /**
     * 将Office文档转换为PDF.
     * <p>
     * 运行该函数需要用到OpenOffice,
     */
    @Test
    public void testOpenOffice() {
        String source = "d:\\tmp\\p2.pptx";
        String target = "d:\\tmp\\pp2.pdf";

        long start = System.currentTimeMillis();

        // 获取系统安装的OpenOffice
        OfficeManager officeManager = LocalOfficeManager.install();

        try {
            // 找不到源文件, 则返回false
            File inputFile = new File(source);

            // 已经存在pdf文件，删除
            File outputFile = new File(target);
            if (outputFile.exists()) {
                outputFile.delete();
            }

            // 启动OpenOffice的服务
            System.out.println("启动OpenOffice...");
            officeManager.start();

            // 转换pdf
            System.out.println("开始使用OpenOffice转换...");
            JodConverter.convert(inputFile).to(outputFile).execute();
            System.out.println("转换完成，耗时：" + (System.currentTimeMillis() - start) + "ms");
        } catch (OfficeException e) {
            e.printStackTrace();
        } finally {
            // 关闭OpenOffice
            OfficeUtils.stopQuietly(officeManager);
        }
    }
}

```



### 3. 参考

- https://github.com/documents4j/documents4j
- https://github.com/sbraconnier/jodconverter
- https://github.com/microacup/office2pdf



