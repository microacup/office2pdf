import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import org.junit.Test;

import java.io.File;

/**
 * 2018/7/13
 */
public class JacobTest {

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
