package com.pojo.poi.test.service;

import com.pojo.poi.test.dto.report.Sales;
import com.pojo.poi.test.dto.report.SalesCategory;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

@Service
public class ReportService {
    public Sales getSalesData() {
        Sales sales = new Sales();
        sales.setYear(2021);
        List<SalesCategory> categories = new ArrayList<>();
        sales.setCategories(categories);
        List<SalesCategory> totalCategories = new ArrayList<>();
        sales.setTotalCategories(totalCategories);
        SalesCategory category1 = new SalesCategory();
        category1.setGubun("직영대리점");
        category1.setMonth1(4582600);
        category1.setMonth2(3582600);
        category1.setMonth3(2582600);
        category1.setMonth4(1582600);
        category1.setMonth5(582600);
        category1.setMonth6(-417400);
        category1.setTotal(total(category1));
        categories.add(category1);
        SalesCategory category2 = new SalesCategory();
        category2.setGubun("기타");
        category2.setMonth1(285000);
        category2.setMonth2(159000);
        category2.setMonth3(33000);
        category2.setMonth4(500000);
        category2.setMonth5(230000);
        category2.setMonth6(170000);
        category2.setTotal(total(category2));
        categories.add(category2);
        SalesCategory category3 = new SalesCategory();
        category3.setGubun("노트");
        category3.setMonth1(445000);
        category3.setMonth2(390000);
        category3.setMonth3(335000);
        category3.setMonth4(280000);
        category3.setMonth5(225000);
        category3.setMonth6(170000);
        category3.setTotal(total(category3));
        categories.add(category3);
        SalesCategory category4 = new SalesCategory();
        category4.setGubun("복사용지");
        category4.setMonth1(13300);
        category4.setMonth2(50000);
        category4.setMonth3(86700);
        category4.setMonth4(123400);
        category4.setMonth5(160100);
        category4.setMonth6(196800);
        category4.setTotal(total(category4));
        categories.add(category4);
        SalesCategory category5 = new SalesCategory();
        category5.setGubun("필기구");
        category5.setMonth1(152600);
        category5.setMonth2(100000);
        category5.setMonth3(47400);
        category5.setMonth4(55000);
        category5.setMonth5(60000);
        category5.setMonth6(74000);
        category5.setTotal(total(category5));
        categories.add(category5);
        SalesCategory category6 = new SalesCategory();
        category6.setGubun("가맹대리점");
        category6.setMonth1(10100);
        category6.setMonth2(22000);
        category6.setMonth3(33900);
        category6.setMonth4(45800);
        category6.setMonth5(57700);
        category6.setMonth6(69600);
        category6.setTotal(total(category6));
        categories.add(category6);
        SalesCategory category7 = new SalesCategory();
        category7.setGubun("총합계");
        category7.setMonth1(sumMonth(categories, SalesCategory::getMonth1));
        category7.setMonth2(sumMonth(categories, SalesCategory::getMonth2));
        category7.setMonth3(sumMonth(categories, SalesCategory::getMonth3));
        category7.setMonth4(sumMonth(categories, SalesCategory::getMonth4));
        category7.setMonth5(sumMonth(categories, SalesCategory::getMonth5));
        category7.setMonth6(sumMonth(categories, SalesCategory::getMonth6));
        category7.setTotal(total(category7));
        totalCategories.add(category7);

        return sales;
    }

    private int sumMonth(List<SalesCategory> categories, Function<SalesCategory, Integer> indentify) {
        return categories.stream().map(indentify).reduce(0, Integer::sum);
    }

    private int total(SalesCategory salesCategory) {
        return salesCategory.getMonth1() + salesCategory.getMonth2() + salesCategory.getMonth3() + salesCategory.getMonth4() + salesCategory.getMonth5() + salesCategory.getMonth6();
    }
}
