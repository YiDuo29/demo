package com.example.demo;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class TextToWordWithName {

    private static final String INPUT_DIR = "C:/Users/17791/Desktop/PrivateSportsField/PrivateSportsField-Frontend/src/assets"; // 替换为您的输入目录
    private static final String OUTPUT_FILE = "C:/Users/17791/Desktop/PrivateSportsField/PrivateSportsField-Frontend/assets.docx";

    public static void main(String[] args) throws IOException, InterruptedException {
        ExecutorService executorService = Executors.newCachedThreadPool();
        XWPFDocument document = new XWPFDocument();

        Files.walkFileTree(Paths.get(INPUT_DIR), new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                if (!attrs.isDirectory()) {
                    executorService.submit(() -> {
                        try {
                            //Q:换行符是什么？
                            //A:换行符是"\n"，回车符是"\r"，Windows系统里，每行结尾是“”，Unix系统里是“\n”，Mac系统里是“\r”。

                            String content = new String(file.getFileName() + "/r" + file + "/r").trim();
                            content += new String(Files.readAllBytes(file));
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

