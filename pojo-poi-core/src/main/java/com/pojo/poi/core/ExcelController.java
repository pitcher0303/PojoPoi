package com.pojo.poi.core;

import com.pojo.poi.core.excel.ExcelMaster;
import com.pojo.poi.core.excel.ExcelModel;
import com.pojo.poi.core.sample.Category;
import com.pojo.poi.core.sample.Project;
import com.pojo.poi.core.sample.Report;
import lombok.SneakyThrows;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

@RestController("/api/v1/excel")
public class ExcelController {
    @RequestMapping(value = "/download/excel", method = RequestMethod.GET, produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<InputStreamResource> downloadExcel() {
        Report report = sampleReport();
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

    @SneakyThrows
    @RequestMapping(value = "/upload/excel", method = RequestMethod.POST, consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<Report> uploadExcel(@RequestPart(value = "file") MultipartFile multipartFile) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(multipartFile.getInputStream()));
        Report report = ExcelMaster.readSheet(Report.class, workbook.getSheetAt(0));
        return ResponseEntity.ok(report);
    }

    public Report sampleReport() {
        Report report = new Report();
        report.getProjectList().add(sampleProject1());
        report.getProjectList().add(sampleProject2());

        return report;
    }

    public Project sampleProject1() {
        Project project = new Project();
        project.setProjectType("SI");
        project.setProjectName("NW빅데이터 분석시스템 \nSSO적용 및 통합 권한 관리 체계 구축");
        project.setProjectManager("박현찬");
        project.setProgressRate("95%");
        project.setIssues("""
                금주, 차주 구분이 모호 하여 업무 처리 대상을 밑줄 표기함.
                1. 운영기 SSL 적용 지연
                 - 스위치 포트바인딩이 지연중.
                2. 보안 검수 진행 중
                """);

        Category category1 = new Category();
        project.getCategories().add(category1);
        category1.setCategoryType("1. 관리");
        category1.setThisWeek("");
        category1.setNextWeek("");

        Category category2 = new Category();
        project.getCategories().add(category2);
        category2.setCategoryType("2. 분석");
        category2.setThisWeek("");
        category2.setNextWeek("");

        Category category3 = new Category();
        project.getCategories().add(category3);
        category3.setCategoryType("3. 설계");
        category3.setThisWeek("");
        category3.setNextWeek("");

        Category category4 = new Category();
        project.getCategories().add(category4);
        category4.setCategoryType("4. 개발(환경 구축)");
        category4.setThisWeek("""
                . Nudge 환경 구축(100%)
                 - 업무, 보안 용 두가지 서버 구축(100%)
                 - 운영기 환경에서 테스트 진행(101%)
                  * 스케줄링 작업의 동시성 Lock 기능 테스트
                  * 세션 클러스터링 테스트
                  * WebSoscket 클러스터링 테스트
                  * 업무/보안 Web 화면간 메뉴 구성 테스트
                """);
        category4.setNextWeek("""
                . Nudge 환경 구축(100%)
                 - 업무, 보안 용 두가지 서버 구축(100%)
                 - 운영기 환경에서 테스트 진행(101%)
                  * 스케줄링 작업의 동시성 Lock 기능 테스트
                  * 세션 클러스터링 테스트
                  * WebSoscket 클러스터링 테스트
                  * 업무/보안 Web 화면간 메뉴 구성 테스트
                """);

        Category category5 = new Category();
        project.getCategories().add(category5);
        category5.setCategoryType("4. 개발(Nudge-1)");
        category5.setThisWeek("""
                . Nudge API 개발(100%)
                  - 시스템 권한 API(100%)
                  - 기존 API 고도화(50%)
                   * 연관관계 효율적 설정, 중복 API 통합(50%)
                  - 승인함 조회 API 개선(100%)
                  - 메뉴 권한 - API 간 권한 적용(100%)
                  - 메뉴 권한 API 개선(100%)
                  - 메뉴 조회 API 개선(100%)
                  - 사용 가능 업무 시스템 API 개발(100%)
                
                . Nudge Web 개발(100%)
                  - 메뉴 조회 API 개선(100%)
                  - 사용가능 시스템 기능 개발(100%)
                  - Nudge 권한 없는 사용자 처리 기능 개발(100%)
                  - 결재 API 리팩토링(100%) *
                
                . Nudge 마당 SSO 연동(100%)
                  - 마당 SSO 연동(100%)
                  - 마당 SSO 연동시 등록되지 않은 사용자 처리(100%)
                  - OAuth 연동시 미등록 사용자 처리 기능 개발(100%)
                """);
        category5.setNextWeek("""
                . Nudge API 개발(100%)
                  - 시스템 권한 API(100%)
                  - 기존 API 고도화(50%)
                   * 연관관계 효율적 설정, 중복 API 통합(50%)
                  - 승인함 조회 API 개선(100%)
                  - 메뉴 권한 - API 간 권한 적용(100%)
                  - 메뉴 권한 API 개선(100%)
                  - 메뉴 조회 API 개선(100%)
                  - 사용 가능 업무 시스템 API 개발(100%)
                
                . Nudge Web 개발(100%)
                  - 메뉴 조회 API 개선(100%)
                  - 사용가능 시스템 기능 개발(100%)
                  - Nudge 권한 없는 사용자 처리 기능 개발(100%)
                  - 결재 API 리팩토링(100%)
                
                . Nudge 마당 SSO 연동(100%)
                  - 마당 SSO 연동(100%)
                  - 마당 SSO 연동시 등록되지 않은 사용자 처리(100%)
                  - OAuth 연동시 미등록 사용자 처리 기능 개발(100%)
                """);

        Category category6 = new Category();
        project.getCategories().add(category6);
        category6.setCategoryType("4. 개발(Nudge-2)");
        category6.setThisWeek("""
                . Nudge Schedule 고도화(100%)
                  - Nudge Schedule 동작시 Multi Server 잠금 기능(100%)
                
                . Nudge Websocket 서버 구축(100%)
                  - resource upload API 개발(100%)
                  - nginx resource download 개발(100%)
                  - 운영기 nginx resource server 구축(100%)
                """);
        category6.setNextWeek("""
                . Nudge Schedule 고도화(100%)
                  - Nudge Schedule 동작시 Multi Server 잠금 기능(100%)
                
                . Nudge Websocket 서버 구축(100%)
                  - resource upload API 개발(100%)
                  - nginx resource download 개발(100%)
                  - 운영기 nginx resource server 구축(100%)
                """);

        Category category7 = new Category();
        project.getCategories().add(category7);
        category7.setCategoryType("4. 개발(CI/CD)");
        category7.setThisWeek("""
                . Nudge 배포 환경 구축(100%)
                  - TeamCity 를 활용한 Spring 기반 프로세스 배포 작업 구성(100%)
                  - TeamCity 를 활용한 Nodejs, Yarn 기반 프로세스 배포 작업 구성(100%)
                
                . Nudge 배포 환경 테스트(100%)
                  - Nudge MSA 배포 테스트(100%)
                  - Nudge MSA 일괄 중지/시작 테스트(100%)
                """);
        category7.setNextWeek("""
                . Nudge 배포 환경 구축(100%)
                  - TeamCity 를 활용한 Spring 기반 프로세스 배포 작업 구성(100%)
                  - TeamCity 를 활용한 Nodejs, Yarn 기반 프로세스 배포 작업 구성(100%)
                
                . Nudge 배포 환경 테스트(100%)
                  - Nudge MSA 배포 테스트(100%)
                  - Nudge MSA 일괄 중지/시작 테스트(100%)
                """);

        Category category8 = new Category();
        project.getCategories().add(category8);
        category8.setCategoryType("5. 검수");
        category8.setThisWeek("""
                 - 연동 시스템별 권한 관리 방안 문서 작성
                 - 프로세스 설계서 작성
                . 보안 검수(50%)
                 - 보안 검수 수정 사항 수정
                  * 검수 진행중
                . 산출물 작성(100%)
                 - 프로젝트 구조도
                 - TeamCity 배포 매뉴얼
                 - 프로세스 상세 설계서
                """);
        category8.setNextWeek("""
                 - 연동 시스템별 권한 관리 방안 문서 작성
                 - 프로세스 설계서 작성
                . 보안 검수(50%)
                 - 보안 검수 수정 사항 수정
                  * 검수 진행중
                . 산출물 작성(100%)
                 - 프로젝트 구조도
                 - TeamCity 배포 매뉴얼
                 - 프로세스 상세 설계서
                """);

        return project;
    }
    public Project sampleProject2() {
        Project project = new Project();
        project.setProjectType("SI");
        project.setProjectName("NW빅데이터 분석시스템 \nDenodo 연동");
        project.setProjectManager("박현찬");
        project.setProgressRate("10%");
        project.setIssues("""
                금주, 차주 구분이 모호 하여 업무 처리 대상을 밑줄 표기함.
                """);

        Category category1 = new Category();
        project.getCategories().add(category1);
        category1.setCategoryType("1. 관리");
        category1.setThisWeek("");
        category1.setNextWeek("");

        Category category2 = new Category();
        project.getCategories().add(category2);
        category2.setCategoryType("2. 분석");
        category2.setThisWeek("""
                . Denodo Okta, 권한 연동 방안(20%)
                  - Denodo View 권한, 개인정보 조회 권한연동을 위한 분석
                   * (View 권한) 데이터 카탈로그 API 분석
                   * (View 권한) View 권한 연동 인터페이스 및 연동 프로시저 분석
                """);
        category2.setNextWeek("""
                . Denodo Okta, 권한 연동 방안(20%)
                  - Denodo View 권한, 개인정보 조회 권한연동을 위한 분석
                   * (View 권한) 데이터 카탈로그 API 분석
                   * (View 권한) View 권한 연동 인터페이스 및 연동 프로시저 분석
                """);

        Category category3 = new Category();
        project.getCategories().add(category3);
        category3.setCategoryType("3. 설계");
        category3.setThisWeek("");
        category3.setNextWeek("");

        Category category4 = new Category();
        project.getCategories().add(category4);
        category4.setCategoryType("4. 개발");
        category4.setThisWeek("""
                . Denodo Okta, 권한 연동 방안(60%)
                  - Denodo 그룹 생성을 위한 테이블 설계(100%)
                  - Denodo 그룹 생성을 위한 테이블 스키마 변경(100%)
                  - Denodo 그룹 생성을 위한 테이블 스키마 추가 변경(100%)
                """);
        category4.setNextWeek("""
                . Denodo Okta, 권한 연동 방안(60%)
                  - Denodo 그룹 생성을 위한 테이블 설계(100%)
                  - Denodo 그룹 생성을 위한 테이블 스키마 변경(100%)
                  - Denodo 그룹 생성을 위한 테이블 스키마 추가 변경(100%)
                  - Denodo View 권한 연동 인터페이스 테이블 설계(100%)
                  - Denodo 개인정보 권한 연동 인터페이스 테이블 설계(100%)
                
                . Denodo 인터페이스 배치 개발(100%)
                  - Denodo 그룹 생성을 위한 사용자 정보 데이터 동기화 배치 개발(100%)
                   * 기존에 있던 user-sync-batch 모듈에 기능 추가 예정
                
                . Denodo 인터페이스 배치 개발(60%)
                  - Denodo 그룹 생성을 위한 사용자 정보 데이터 동기화 배치 개발(100%)
                
                """);

        Category category7 = new Category();
        project.getCategories().add(category7);
        category7.setCategoryType("4. 개발(CI/CD)");
        category7.setThisWeek("""
                """);
        category7.setNextWeek("""
                """);

        Category category8 = new Category();
        project.getCategories().add(category8);
        category8.setCategoryType("5. 검수");
        category8.setThisWeek("""
                """);
        category8.setNextWeek("""
                """);

        return project;
    }

}
