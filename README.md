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

### 프로젝트 구성

1. __pojo-poi-core__ : PojoPoi 핵심 모듈
2. __pojo-poi-test__ : PojoPoi 를 Dependency 하여 테스트 하는 샘플 코드

### PojoPoi Design

이 라이브러리는 Excel 을 Java Object 로 쉽게 만들기 위해 제작 되었습니다.

Excel 의 형태를 Java Object, Meta Data (Annotation 사용) 으로 디자인하고 즉시 Excel 을 만들거나 읽어 들일 수 있습니다.

1. 엑셀 형식을 정합니다. 예시에서는 아래 형태 엑셀을 만들고, 읽을 것입니다.

![Sample1-01](https://github.com/user-attachments/assets/ae9ca41e-9817-4f17-8e24-e9798b6abfa4)

2. 엑셀의 sheet 내 구역을 Java Object 에 맞춰 디자인 합니다.

![Sample1-02](https://github.com/user-attachments/assets/fce69e2c-0c3a-412b-a0fe-5bc6fd106e13)

* 상반기 월별 매출 현황 을 _Sales.java_ 로 정합니다.
* B3 ~ B9 구역은 형태가 같으므로 RowMeta(List 타입 처리)로 정합니다.
* B3 구역은 RowMeta 의 HeaderMeta 를 사용합니다.
* B4 ~ B9 구역을 Row 한줄당 _SalesCategory.java_ 로 정합니다.
* B10 구역은 동일한 _SalesCategory.java_ 형태이나 예시를 위해 따로 매핑합니다.
* 위 상황에 맞춰 아래 처럼 적절한 Annotation 으로 매핑해 주면 됩니다.
* ~~각 Annotation 에 대한 설명은 추후 업로드 예정~~

![Sample1-03](https://github.com/user-attachments/assets/2b3fb90a-26bc-4a25-a7b5-da3a65309637)

![Sample1-04](https://github.com/user-attachments/assets/2238afc5-9af4-456b-81b2-65638fdc575b)

* 위와 같이 _Sales.java_, _SalesCategory.java_ 두 가지 객체만 사용하여 엑셀을 디자인하고 Java Object 로 변환 합니다.

3.  Read, Write 로 Excel 을 읽거나, 쓸수 있습니다.

* 액셀을 쓰거나 읽을 시 별도 객체 생성, 설정 필요 없이 Sales 자체로 읽고, 쓸수 있습니다.

Write Excel

```java
//Sales 데이터를 비즈니스 로직에서 가져오기
Sales sales = reportService.getSalesData();
//ExcelWriter 를 사용하여 (Sheet 명, 데이터, Column With) 를 전달하고 쓰기
ExcelWriter excelWriter = ExcelWriter.builder("샘플1")
        .build()
        .addExcelDatas("샘플1", List.of(sales), new float[]{8.38f, 21, 11, 11, 11, 11, 11, 11, 15})
        .writeAll();
//ExcelWriter 로 InputStream 가져오기
InputStreamResource resource = new InputStreamResource(excelWriter.getExcelStream());
//... MediaType.APPLICATION_OCTET_STREAM_VALUE 으로 반환하기 ...
//... 생략 ...
```

Read Excel

```java
@SneakyThrows
@RequestMapping(value = "/sales/upload/excel", method = RequestMethod.POST, consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
//MultipartFile 로 Excel 받기
public ResponseEntity<Sales> salesUploadExcel(@RequestPart(value = "file") MultipartFile multipartFile) {
    //Poi Workbook 객체 가져오기
    XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(multipartFile.getInputStream()));
    //ExcelReader 로 Sales 객체로 변환 시키기
    Sales sales = ExcelReader.readSheet(Sales.class, workbook.getSheetAt(0));
    return ResponseEntity.ok(sales);
}
```