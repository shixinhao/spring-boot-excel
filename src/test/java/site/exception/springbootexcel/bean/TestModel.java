package site.exception.springbootexcel.bean;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

public class TestModel extends BaseRowModel {

    @ExcelProperty(value = "姓名", index = 0)
    private String xm;
    @ExcelProperty(value = "微信号", index = 1)
    private String wxh;
    @ExcelProperty(value = "手机号", index = 2)
    private String sjh;
    @ExcelProperty(value = "haha", index = 3)
    private String haha;


    public String getHaha() {
        return haha;
    }

    public void setHaha(String haha) {
        this.haha = haha;
    }

    public String getXm() {
        return xm;
    }
    public void setXm(String xm) {
        this.xm = xm;
    }
    public String getWxh() {
        return wxh;
    }
    public void setWxh(String wxh) {
        this.wxh = wxh;
    }
    public String getSjh() {
        return sjh;
    }
    public void setSjh(String sjh) {
        this.sjh = sjh;
    }

}
