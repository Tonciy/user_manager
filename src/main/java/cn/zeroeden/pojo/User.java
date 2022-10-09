package cn.zeroeden.pojo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;
import tk.mybatis.mapper.annotation.KeySql;

import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;
import java.util.List;

/**
 * 员工
 */
@Data
@Table(name = "tb_user")
public class User {
    @Id
    @KeySql(useGeneratedKeys = true)
    @Excel(name = "编号", orderNum = "1", width = 5)
    private Long id;         //主键
    @Excel(name = "姓名", orderNum = "2", width = 15)
    private String userName; //员工名
    @Excel(name = "手机号", orderNum = "3", width = 15)
    private String phone;    //手机号
    @Excel(name = "省份", orderNum = "4", width = 15)
    private String province; //省份名
    @Excel(name = "城市", orderNum = "5", width = 15)
    private String city;     //城市名
    @Excel(name = "工资", orderNum = "5", width = 15, type = 10)
    private Integer salary;   // 工资
    @Excel(name = "入职日期", orderNum = "6", width = 15, format = "yyyy-MM-dd")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date hireDate; // 入职日期
    private String deptId;   //部门id
    @Excel(name = "出生日期", orderNum = "7", width = 15, format = "yyyy-MM-dd")
    private Date birthday; //出生日期
    @Excel(name = "照片", orderNum = "9", type = 2)
    private String photo;    //一寸照片
    @Excel(name = "现住址", orderNum = "8", width = 20)
    private String address;  //现在居住地址

    private List<Resource> resourceList; //办公用品

}
