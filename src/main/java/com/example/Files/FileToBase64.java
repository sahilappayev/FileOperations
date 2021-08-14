package com.example.Files;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Base64;

@RestController
@RequestMapping("/files")
public class FileToBase64 {


    @GetMapping
    public ResponseEntity<String> getBase64() {
        String directory = "C:\\Users\\AppayevS\\Desktop\\videoplayback.mp4";
        try {
            File file = new File(directory);
            byte[] bytes = Files.readAllBytes(file.toPath());
            String encoded = new String(Base64.getEncoder().encode(bytes));
            return ResponseEntity.ok(encoded);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

}
