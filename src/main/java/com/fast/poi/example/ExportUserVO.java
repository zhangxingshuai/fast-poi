package com.fast.poi.example;

import com.fast.poi.excel.annotation.ExcelDropDownBox;
import com.fast.poi.excel.annotation.ExcelField;

/**
 * 导出用户测试类 数据位置： /resource/mock/user.txt
 *
 * @author 张兴帅
 * @projectName aidustry-fast
 * @packageNmae com.fast.test.vo
 * @fileNmae ExportUserVO
 * @date 2023-9-1
 * @copyright 
 * @since 0.0.1
 */

public class ExportUserVO {

    @ExcelField(title = "姓名")
    private String name;

    @ExcelField(title = "性别")
    @ExcelDropDownBox(values = {"男","女"})
    private String gender;

    @ExcelField(title = "年龄")
    private Integer age;

    @ExcelField(title = "城市")
    private String city;

    @ExcelField(title = "政治面貌")
    @ExcelDropDownBox(values = {"党员","团员","群众","无党派人士"}, valuesMap = {"1","2","3","4"})
    private Integer political;

    public Integer getPolitical() {
        return political;
    }

    public void setPolitical(Integer political) {
        this.political = political;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }
}
