<template>
  <div class="index" v-loading.fullscreen.lock="fullscreenLoading" element-loading-text="拼命加载中...">
    <input
      type="file"
      @change="importFile(this)"
      id="imFile"
      style="display: none"
      accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
    />
    <a id="downlink"></a>
    <el-button class="button" @click="uploadFile()">数据导入</el-button>
    <el-button class="button" @click="downloadFile(multipleSelection)">单据批量生成</el-button>
    <!--错误信息提示-->
    <el-dialog title="提示" v-model="errorDialog" size="tiny">
      <span>{{errorMsg}}</span>
      <span slot="footer" class="dialog-footer">
        <el-button type="primary" @click="errorDialog=false">确认</el-button>
      </span>
    </el-dialog>
    <!--展示导入信息-->
    <el-table
      align="center"
      :data="excelendData"
      tooltip-effect="dark"
      @selection-change="handleSelectionChange"
    >
      <el-table-column type="selection" width="55"></el-table-column>
      <el-table-column label="用户编号" prop="id" show-overflow-tooltip></el-table-column>
      <el-table-column label="用户名称" prop="name" show-overflow-tooltip></el-table-column>
      <el-table-column label="用电地址" prop="addr" show-overflow-tooltip></el-table-column>
      <el-table-column label="基本电费" prop="basicElectricity" show-overflow-tooltip></el-table-column>
      <el-table-column label="电压等级" prop="Voltage" show-overflow-tooltip></el-table-column>
      <el-table-column label="电表局编号" prop="linename" show-overflow-tooltip></el-table-column>
      <el-table-column label="计量点电价" prop="prices" show-overflow-tooltip></el-table-column>
      <el-table-column label="发现时间" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.time}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="合同容量" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.capacity}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="月运行容量" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.moncapacity}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="超容比" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.supercapacity}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="实际使用容量" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.actualcapacity}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="补缴基本电费" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.paymentele}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column label="违约使用电费" show-overflow-tooltip>
        <template slot-scope="scope">
          <!-- {{scope.row.children}} -->
          <template v-for="item in scope.row.children">
            <tr>
              <td>{{item.defaulttele}}</td>
            </tr>
          </template>
        </template>
      </el-table-column>
      <el-table-column fixed="right" label="操作" width="100">
        <template slot-scope="scope">
          <el-button @click="downloadExlbyonly(scope.row)" size="mini">单据导出</el-button>
        </template>
      </el-table-column>
    </el-table>
  </div>
</template>

<script>
// 引入xlsx
import base64 from "js-base64";
var XLSX = require("xlsx");
export default {
  name: "Index",
  data() {
    return {
      fullscreenLoading: false, // 加载中
      imFile: "", // 导入文件el
      outFile: "", // 导出文件el
      errorDialog: false, // 错误信息弹窗
      errorMsg: "", // 错误信息内容
      excelData: [
        // 初步处理数据
      ],
      excelendData: [
        // 二次处理数据可输出数据
      ],
      multipleSelection: [], //已选中处理项
      excelDatajson: {
        id: "",
        name: "",
        addr: "",
        basicElectricity: "", //基本电费
        Voltage: "", //电压等级
        moncapacity: "", //月运行容量
        supercapacity: "", //超容比
        linename: "", //线路名称
        prices: "", //电价
        time: "" // 时间
      },
      Datajson: {
        id: "",
        name: "",
        addr: "",
        capacity: "", //合同容量
        basicElectricity: "", //基本电费
        Voltage: "", //电压等级
        linename: "", //线路名称
        prices: "", //电价
        children: [
          {
            moncapacity: null, //月运行容量
            actualcapacity: null, //实际使用容量
            supercapacity: null, //超容比
            paymentele: null, //补缴基本电费
            defaulttele: null, //违约补缴电费
            time: null // 时间
          }
        ]
      }
    };
  },
  mounted() {
    this.imFile = document.getElementById("imFile");
    this.outFile = document.getElementById("downlink");
  },
  methods: {
    handleSelectionChange(val) {
      this.multipleSelection = val;
      // console.log("已选中处理项", this.multipleSelection);
    },
    uploadFile: function() {
      // 去掉表头导入
      this.imFile.click();
    },
    // 单对象文件导出
    downloadExlbyonly(data) {
      // console.log("单对象导出", data);
      let arrdata = [];
      arrdata.push(data);
      this.downloadFile(arrdata);
    },
    //文件导出
    downloadFile: function(excelobjarr) {
      // 点击导出按钮
      if (excelobjarr.length <= 0) {
        this.errorDialog = false;
        this.errorMsg = "未选择输出表格对象，请选中后操作";
      } else {
        for (const excelitem of excelobjarr) {
          let excelname = excelitem.name + "_超额电量罚单";
          this.downloadExl(excelitem, excelname);
        }
      }
    },
    downloadExl: function(json, downName, type) {
      // 导出到excel
      let worksheet = "Sheet1";
      let id = json.id;
      let name = json.name;
      let addr = json.addr;
      let str = "";
      //循环遍历，每行加入tr标签，每个单元格加td标签
      let acmu = 0; //实际用量和
      // paymentele: null, //补缴基本电费
      //     defaulttele: null, //违约补缴电费
      let paymentelemum = 0; //补缴基本电费和
      let defaulttelemum = 0; //违约电费和
      let Subtotalcombined = 0; // 小计合计
      let heightcol = 7 + json.children.length + 1;
      for (const item of json.children) {
        let submoney = item.defaulttele + item.paymentele; //单月电费小计
        Subtotalcombined = Subtotalcombined + submoney; //单对象电费累计
        acmu = acmu + item.actualcapacity; //实际使用量合计
        paymentelemum = paymentelemum + item.paymentele; //补缴基本电费和
        defaulttelemum = defaulttelemum + item.defaulttele; //违约电费和
        str += "<tr style=" + "color:red;text-align:center;" + ">";
        str += `<td>${item.time}</td>`;
        str += `<td>${item.capacity}</td>`;
        str += `<td>${item.actualcapacity}</td>`;
        str += `<td>${item.paymentele}</td>`;
        str += `<td>${item.defaulttele}</td>`;
        str += `<td>${submoney}</td>`;
        str += "</tr>";
      }
      str += "<tr style=" + "text-align:center;" + ">";
      str += `<td>合计</td>`;
      str += `<td>&nbsp;</td>`;
      str += `<td>${acmu}</td>`;
      str += `<td>${paymentelemum}</td>`;
      str += `<td>${defaulttelemum}</td>`;
      str += `<td>${Subtotalcombined}</td>`;
      str += "</tr>";
      //Worksheet名

      let template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
      xmlns:x="urn:schemas-microsoft-com:office:excel" 
      xmlns="http://www.w3.org/TR/REC-html40">
      <head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
        <x:Name>${worksheet}</x:Name>
        <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
        </x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
        </head><body>
        <table border="1" cellspacing="0"  width="700" style="font-size: 14px;border-spacing: 0px;border-collapse:collapse;text-align:center;">
            <tr>
              <th width="50">户名</th>
              <th width="140" style="color:red">${name}</th>
              <th width="80">户号</th>
              <th width="100" style="color:red">${id}</th>
              <th width="100">地址</th>
              <th colspan="2" width="200" style="color:red">${addr}</th>
            </tr>
            <tr>
              <td rowspan="${heightcol}"></td>
              <td
                colspan="6"
                style="font-size: 16px;text-align: left;"
              >该户合同容量为250kVA，擅自超过合同约定容量用电，根据《电力法》 第四十条规定： 违反本条例第三十条规定，违章用电的，供电企业可以根据违章事实和造成的后果追缴电费，并按照国务院电力管理部门的规定加收电费和国家规定的其他费用；情节严重的，可以按照国家规定的程序停止供电。现根据《供电营业规则》第一百条规定：擅自超过本合同约定容量用电的，属于两部制电价的用户，按三倍私增容量基本电费计付违约使用电费；属单一制电价的用户，按擅自使用或启封设备容量每千伏安/千瓦50元支付违约使用电费；现补缴基本电费及违约使用电费，计算如下表，并请办理增容手续。</td>
            </tr>
            <tr height='60' style="text-align: center;">
              <td   >使用月份</td>
              <td  >合同容量（kVA）</td>
              <td  >实际使用容量（kVA）</td>
              <td  >补缴基本电费（元）</td>
              <td  >违约使用电费（元）</td>
              <td  > 小计（元）</td>
            </tr>
            ${str}
            <tr>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td>当事人</td>
              <td style="border-right: 0px solid black;">签章</td>
              <td style="border-left: 0px solid black;">&nbsp;</td>
              
            </tr>
            <tr style="text-align: center;">
              <td  style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td>检测人</td>
              <td style="border-right: 0px solid black;">签章</td>
              <td style="border-left: 0px solid black;">&nbsp;</td>
              
            </tr>
            <tr style="text-align: center;">
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td>处理人</td>
              <td style="border-right: 0px solid black;">签章</td>
              <td style="border-left: 0px solid black;">&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td style="border-top: 0px solid black;border-left: 0px solid black;border-right: 0px solid black;">&nbsp;</td>
              <td style="border-top: 0px solid black;border-left: 0px solid black;border-right: 0px solid black;">&nbsp;</td>
              <td style="border-top: 0px solid black;border-left: 0px solid black;border-right: 0px solid black;">&nbsp;</td>
              <td>年</td>
              <td>月</td>
              <td>日</td>
            </tr>
            <tr>
              <td rowspan="7"></td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr >
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">负责人（签章）</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">年&nbsp;&nbsp;月</td>
              <td style="border: 0px solid black;">日</td>
            </tr>
            <tr>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
              <td style="border: 0px solid black;">&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td rowspan="3">收费
                <br/>
                记录</td>
              <td>项目</td>
              <td>数额</td>
              <td colspan="2">收据号</td>
              <td>收款人</td>
              <td>&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td>补电费(电量)</td>
              <td>&nbsp;</td>
              <td colspan="2">&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td>违约使用电费</td>
              <td>&nbsp;</td>
              <td colspan="2">&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr style="text-align: center;">
              <td>业务</td>
              <td>&nbsp;</td>
              <td>业务记录</td>
              <td colspan="2">&nbsp;</td>
              <td>电费记录</td>
              <td>&nbsp;</td>
            </tr>
          </table>
        </body></html>`;
      //下载模板
      let uri = "data:application/vnd.ms-excel;base64,";
      // window.location.href = uri + this.base64(template)
      var href = uri + this.base64(template); // 创建对象超链接
      // console.log(href, "表格文件");
      // var href = URL.createObjectURL(tmpDown); // 创建对象超链接
      this.outFile.download = downName + ".xls"; // 下载名称
      this.outFile.href = href; // 绑定a标签
      this.outFile.click(); // 模拟点击实现下载
      // setTimeout(function() {
      //   // 延时释放
      
      //   URL.revokeObjectURL(tmpDown); // 用URL.revokeObjectURL()来释放这个object URL
      // }, 100);
    },
    base64(s) {
      return window.btoa(unescape(encodeURIComponent(s)));
    },
    importFile: function() {
      // 导入excel解析为json
      this.fullscreenLoading = true;
      let obj = this.imFile; //表格对象
      // console.log("导入表格", obj);
      if (!obj.files) {
        this.fullscreenLoading = false;
        return;
      }
      var f = obj.files[0];
      var reader = new FileReader();
      let $t = this;
      reader.onload = function(e) {
        var data = e.target.result; //导入表格信息
        //导入表格详细信息转换
        if ($t.rABS) {
          $t.wb = XLSX.read(btoa(this.fixdata(data)), {
            // 手动转化
            type: "base64"
          });
        } else {
          $t.wb = XLSX.read(data, {
            type: "binary"
          });
        }
        //表格信息处理 --第一个工作表内数据
        let json = XLSX.utils.sheet_to_json($t.wb.Sheets[$t.wb.SheetNames[0]]);
        // console.log(typeof json);
        // console.log("处理数据", json); //导入表格数据
        $t.dealFile($t.analyzeData(json)); // analyzeData: 解析导入数据
      };
      if (this.rABS) {
        reader.readAsArrayBuffer(f);
      } else {
        reader.readAsBinaryString(f);
      }
    },
    analyzeData: function(data) {
      // 此处可以解析导入数据 ---转为需要的json对象
      let changedata = [];
      for (const value of data) {
        // console.log(value, "循环处理");
        if (value["__EMPTY"] != "用户名称") {
          let iddata = null;
          let arrdata = [];
          for (const key in value) {
            arrdata.push(value[key]);
          }
          changedata.push({
            id: arrdata[0],
            name: arrdata[1],
            addr: arrdata[2],
            basicElectricity: arrdata[3], //基本电费
            Voltage: arrdata[4], //电压等级
            capacity: arrdata[5], //合同容量
            moncapacity: arrdata[6], //月运行容量
            supercapacity: arrdata[7], //超容比
            linename: arrdata[9], //线路名称
            prices: arrdata[8], //电价
            time: arrdata[17] // 时间
          });
        }
      }
      data = changedata;
      return data;
    },
    dealFile: function(data) {
      // 处理导入的数据
      // console.log(data, "导入数据处理");
      this.imFile.value = "";
      this.fullscreenLoading = false;
      if (data.length <= 0) {
        this.errorDialog = true;
        this.errorMsg = "请导入正确信息";
      } else {
        this.excelData = this.excelData.concat(data); //  用于不同月份数据评价
      }
      this.agindealdata(this.excelData);
    },
    //累加数据重处理
    agindealdata: function(arrdata) {
      // 处理导入的数据
      // console.log(arrdata, "累加数据重处理");
      let newList = [];
      arrdata.forEach(data => {
        let calculateprices = 90;
        let pricesstatus = 30;
        if (data.prices === "普通工业非优待(10kV)") {
          calculateprices = 50;
          pricesstatus = 0;
        } else {
          calculateprices = 90;
          pricesstatus = 30;
        }
        //id轮训并写入
        for (let i = 0; i < newList.length; i++) {
          if (newList[i].id === data.id) {
            //重复时间去除
            for (const item of newList[i].children) {
              if (item.time == data.time) {
                break;
              } else {
                newList[i].children.push({
                  capacity: data.capacity,
                  moncapacity: data.moncapacity, //月运行容量
                  actualcapacity:
                    data.moncapacity * (1 + data.supercapacity / 100), //实际使用容量
                  supercapacity: data.supercapacity, //超容比
                  paymentele:
                    pricesstatus *
                    data.moncapacity *
                    (data.supercapacity / 100), //补缴基本电费
                  defaulttele:
                    ((data.moncapacity * data.supercapacity) / 100) *
                    calculateprices, //违约补缴电费
                  time: data.time // 时间
                });
              }
            }
            return;
          }
        }
        //非重复写入
        newList.push({
          id: data.id,
          name: data.name,
          addr: data.addr,
          basicElectricity: data.basicElectricity, //基本电费
          Voltage: data.Voltage, //电压等级
          linename: data.linename, //线路名称
          prices: data.prices, //电价
          children: [
            {
              capacity: data.capacity,
              moncapacity: data.moncapacity, //月运行容量
              actualcapacity: data.moncapacity * (1 + data.supercapacity / 100), //实际使用容量
              supercapacity: data.supercapacity, //超容比
              paymentele:
                pricesstatus * data.moncapacity * (data.supercapacity / 100), //补缴基本电费
              defaulttele:
                data.moncapacity * (data.supercapacity / 100) * calculateprices, //违约补缴电费
              time: data.time // 时间
            }
          ]
        });
      });
      // console.log(newList);
      this.excelendData = newList;
    },
    s2ab: function(s) {
      // 字符串转字符流
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i !== s.length; ++i) {
        view[i] = s.charCodeAt(i) & 0xff;
      }
      return buf;
    },
    getCharCol: function(n) {
      // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
      let s = "";
      let m = 0;
      while (n > 0) {
        m = (n % 26) + 1;
        s = String.fromCharCode(m + 64) + s;
        n = (n - m) / 26;
      }
      return s;
    },
    fixdata: function(data) {
      // 文件流转BinaryString
      var o = "";
      var l = 0;
      var w = 10240;
      for (; l < data.byteLength / w; ++l) {
        o += String.fromCharCode.apply(
          null,
          new Uint8Array(data.slice(l * w, l * w + w))
        );
      }
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
      return o;
    }
  }
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style>
.el-table th > .cell {
  text-align: center;
}
.button {
  margin-bottom: 20px;
}
.index {
  width: 90%;
  margin: 0 auto;
}
</style>