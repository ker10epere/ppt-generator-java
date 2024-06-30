package com.pushforward.pptgenerator;

import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;

@Component
public class Main implements ApplicationRunner {
    @Value("classpath:ppt-template.pptx")
    Resource pptxTemplateRes;
    @Value("classpath:lyrics.txt")
    Resource lyricsRes;

    @Override
    public void run(ApplicationArguments args) throws Exception {

        String lyrics = new String(lyricsRes.getContentAsByteArray(), StandardCharsets.UTF_8);
        List<String[]> songs = Arrays.stream(lyrics
                        .split("\\n---"))
                .map(song -> {
                    return song.trim()
                            .replace(System.lineSeparator(), "$$") // MAKE ALL LINE BREAKS AS DOUBLE DOLLARS $
                            .replace("$$$$", "##") // TURN 2 LINE BREAKS TO DOUBLE HASHTAGS #
                            .replace("$$", System.lineSeparator()) // TURN 2 DOLLARS TO LINE BREAK
                            .split("##"); // TURN 2 HASHTAG TO LINE BREAK
                })
                .toList();

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptxTemplateRes.getURI().getPath()));
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        XSLFSlideLayout titleLayout = defaultMaster.getLayout("Title");
        XSLFSlideLayout blankLayout = defaultMaster.getLayout("Blank");
        XSLFSlideLayout contentLayout = defaultMaster.getLayout("Content");

        for (String[] song : songs) {
            String titleText = song[0];

            XSLFSlide titleSlide = ppt.createSlide(titleLayout);
            XSLFTextShape titlePlaceholder = titleSlide.getPlaceholder(0);
            titlePlaceholder.setText(titleText);

            for (int i = 1; i < song.length; i++) {
                XSLFSlide contentSlide = ppt.createSlide(contentLayout);
                XSLFTextShape contentPlaceholder = contentSlide.getPlaceholder(0);
                contentPlaceholder.setText(song[i]);
            }
            ppt.createSlide(blankLayout);
        }

        File file = new File("src/main/resources/generated-pptx/sample.pptx");
        file.createNewFile();
        FileOutputStream fos = new FileOutputStream(file);
        ppt.write(fos);
    }
}
