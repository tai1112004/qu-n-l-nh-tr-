package com.cntt.rentalmanagement.services.impl;


import com.cntt.rentalmanagement.config.FileStorageProperties;
import com.cntt.rentalmanagement.exception.FileStorageException;
import com.cntt.rentalmanagement.exception.MyFileNotFoundException;
import com.cntt.rentalmanagement.services.FileStorageService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

@Service
public class FileStorageServiceImpl implements FileStorageService {
    private Path fileStorageLocation;

    public FileStorageServiceImpl() {
        // TODO document why this constructor is empty
    }

    @Autowired
    public void FileStorageService(FileStorageProperties fileStorageProperties) {
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir())
                .toAbsolutePath().normalize();

        try {
            Files.createDirectories(this.fileStorageLocation);
        } catch (Exception ex) {
            throw new FileStorageException("Invalid");
        }
    }

    @Override
    public String storeFile(MultipartFile file) {
        if (file == null || file.isEmpty() || !StringUtils.hasText(file.getOriginalFilename())) {
            throw new FileStorageException("File upload không hợp lệ");
        }

        String fileName = StringUtils.cleanPath(file.getOriginalFilename());

        try {
            if (fileName.contains("..")) {
                throw new FileStorageException("Tên file không hợp lệ");
            }

            Path targetLocation = this.fileStorageLocation.resolve(fileName);
            if (!targetLocation.normalize().startsWith(this.fileStorageLocation)) {
                throw new FileStorageException("Đường dẫn file không hợp lệ");
            }

            Files.copy(file.getInputStream(), targetLocation, StandardCopyOption.REPLACE_EXISTING);

            return fileName;
        } catch (IOException ex) {
            throw new FileStorageException("Không thể lưu file vào: " + this.fileStorageLocation, ex);
        }
    }

    @Override
    public Resource loadFileAsResource(String fileName) {
        try {
            Path filePath = this.fileStorageLocation.resolve(fileName).normalize();
            Resource resource = new UrlResource(filePath.toUri());
            if(resource.exists()) {
                return resource;
            } else {
                throw new MyFileNotFoundException("Not found!!!");
            }
        } catch (MalformedURLException ex) {
            throw new MyFileNotFoundException("Not found!!!");
        }
    }
}
