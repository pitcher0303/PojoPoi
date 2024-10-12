package com.pojo.poi.core;

import com.pojo.poi.core.excel.ExcelModel;
import com.pojo.poi.core.sample.Report;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

@RestController("/api/v1/excel")
public class ExcelController {
    @RequestMapping(value = "/download/excel", method = RequestMethod.GET, produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<InputStreamResource> downloadExcel() {
        Report report = new Report();
        ExcelModel model = ExcelModel.builder("테스트")
                .build()
                .addExcelDatas(List.of(report))
                .writeAll()
                .end();
        InputStreamResource resource = new InputStreamResource(model.getExcelStream());
        HttpHeaders headers = new HttpHeaders();
        String encodedFileName = URLEncoder.encode(model.getFileName(), StandardCharsets.UTF_8).replace("+", "%20");
        headers.setContentDispositionFormData("attachment", encodedFileName);

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

    @RequestMapping(value = "/upload/excel", method = RequestMethod.POST, produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public void uploadExcel() {
    }
}
