import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class TextToWord2 {

    private static final String INPUT_DIR = "D:/ruoyi"; // 替换为您的输入目录
    private static final String OUTPUT_FILE = "D:/ruoyi/output.docx";

    public static void main(String[] args) throws IOException, InterruptedException {
        ExecutorService executorService = Executors.newCachedThreadPool();
        XWPFDocument document = new XWPFDocument();

        Files.walkFileTree(Paths.get(INPUT_DIR), new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                if (!attrs.isDirectory() && file.toString().toLowerCase().endsWith(".txt")) {
                    System.out.println("正在处理文件：" + file);
                    executorService.submit(() -> {
                        try {
                            String content = new String(Files.readAllBytes(file));
                            synchronized (document) {
                                XWPFParagraph paragraph = document.createParagraph();
                                XWPFRun run = paragraph.createRun();
                                run.setText(content);
                            }
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });
                }
                return FileVisitResult.CONTINUE;
            }
        });

        executorService.shutdown();
        executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.MILLISECONDS);

        try (FileOutputStream out = new FileOutputStream(OUTPUT_FILE)) {
            document.write(out);
        }

        document.close();
        System.out.println("文本已成功复制到Word文件：" + OUTPUT_FILE);
    }
}
