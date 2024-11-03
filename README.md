# PojoPoi

## _PojoPoi_ 소개

Apache Poi, Excel <->  Plain Old Java Object Mapper

PojoPoi 는 Apache Poi 라이브러리를 쉽게 사용하기 위한 Wrapper 라이브러리 입니다.

## Usage

Java Object 로 Excel 형태를 디자인하고 Excel Read, Write 에 사용 할 수 있습니다.

### 구성

지원사항: __XssfWorkBook__(지원), __SxssfWorkBook__(지원예정)

Excel 읽기: _ExcelReader_

읽기 샘플

```java
@SneakyThrows
@RequestMapping(value = "/upload/excel", method = RequestMethod.POST, consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
public ResponseEntity<Report> uploadExcel(@RequestPart(value = "file") MultipartFile multipartFile) {
    XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(multipartFile.getInputStream()));
    Report report = ExcelReader.readSheet(Report.class, workbook.getSheetAt(0));
    return ResponseEntity.ok(report);
}
```

Excel 쓰기: _ExcelWriter_

쓰기 샘플

```java
@RequestMapping(value = "/download/excel", method = RequestMethod.GET, produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
public ResponseEntity<InputStreamResource> downloadExcel() {
    Report report = sampleReport();
    ExcelWriter excelWriter = ExcelWriter.builder("테스트")
            .build()
            .addExcelDatas("보고서", List.of(report), new float[]{12.5f, 31.13f, 6.88f, 12f, 16.25f, 68.75f, 68.75f, 71.75f})
            .writeAll();

    InputStreamResource resource = new InputStreamResource(excelWriter.getExcelStream());
    HttpHeaders headers = new HttpHeaders();
    String encodedFileName = URLEncoder.encode(excelWriter.getFileName(), StandardCharsets.UTF_8).replace("+", "%20");
    headers.setContentDispositionFormData("attachment", encodedFileName);

    return ResponseEntity.ok()
            .headers(headers)
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .body(resource);
}
```

### PojoPoi Design

이 라이브러리는 Excel 을 Java Object 로 쉽게 만들기 위해 제작 되었습니다.

1. 엑셀 형식을 정합니다.

--이미지1.(엑셀 샘플), 이미지2.(엑셀 샘플)

2. 엑셀의 sheet 내 구역을 Java Object 에 맞춰 디자인 합니다.

--예시 이미지

3.  Read, Write 로 Excel 을 읽거나, 쓸수 있습니다.

--예시 이미지