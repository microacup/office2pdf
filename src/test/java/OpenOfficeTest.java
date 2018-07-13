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
     * 运行该函数需要用到OpenOffice
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
