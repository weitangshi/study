<template>
  <div class="threeLevelMain">
    <!-- 报价容器 -->
    <div class="quoteContainer">
      <img :src="form.imgUrl" alt="" />
      <!-- 报价顶部信息 -->
      <div class="quote_info clearfix">
        <h3 class="h3_title">报价单</h3>
        <div class="quote_itemBox">
          <div class="quote_item">
            <span class="quote_label">客户名称：</span>
            <p class="quote_p">{{ form.custName }}</p>
          </div>
          <div class="quote_item">
            <span class="quote_label">联系方式：</span>
            <p class="quote_p">{{ form.phoneNumber }}</p>
          </div>
        </div>
        <div class="quote_itemBox">
          <div class="quote_item">
            <span class="quote_label">项目要求：</span>
            <p class="quote_p">{{ form.projectRequirement }}</p>
          </div>
          <div class="quote_item">
            <span class="quote_label">备注：</span>
            <p class="quote_p">{{ form.remark }}</p>
          </div>
        </div>
      </div>
      <!-- 设备选取表格 -->
      <div class="quote_table clearfix">
        <el-table
          border
          class="table_domQuote"
          ref="tableDomQuote"
          size="small"
          :data="tableData"
          style="width: 100%"
        >
          <el-table-column
            prop="number"
            width="80"
            label="序号"
            align="center"
          ></el-table-column>
          <el-table-column
            label="设备名称"
            prop="name"
            align="center"
          ></el-table-column>
          <el-table-column
            label="数量"
            prop="num"
            align="center"
          ></el-table-column>
          <el-table-column
            prop="salePrice"
            label="销售单价"
            align="center"
          ></el-table-column>
          <el-table-column
            prop="saleTotal"
            label="销售合计"
            align="center"
          ></el-table-column>
          <el-table-column
            label="备注"
            prop="remark"
            align="center"
          ></el-table-column>
          <div slot="append">
            <div class="quoteTable">
              <span class="quoteTable_sp1">合计：</span>
              <span class="quoteTable_sp1">{{ form.totalPrice }}</span>
            </div>
          </div>
        </el-table>
      </div>
    </div>
    <!-- 审核备注容器 -->
    <div class="reasonBox clearfix">
      <div class="title_checkReason">审核备注：</div>
      <div class="checkReasonMain">
        <div class="p_box">
          <p class="p_checkReason">{{ form.checkReason }}</p>
        </div>
      </div>
    </div>
    <!-- 底部按钮容器 -->
    <div class="botmBtnContainer">
      <el-button @click="exportWord" size="small" type="primary"
        >导出word</el-button
      >
      <!-- <el-button @click="exportExcelClick" size="small" type="primary">导出excel</el-button> -->
    </div>
  </div>
</template>
<script>
import JSZipUtils from 'jszip-utils'
import docxtemplater from 'docxtemplater'
import { saveAs } from 'file-saver'
import PizZip from 'pizzip'

export default {
  name: "home",
  data () {
    return {
      // 表单对象
      form: {
        custName: "杰斯", // 客户姓名
        phoneNumber: "138xxxxxxxx", // 联系方式
        projectRequirement: "为了更美好的明天而战", // 项目要求
        totalPrice: 140, // 合计报价
        imgUrl: require('../assets/logo.jpg'),
        remark: "QEWARAEQAAAAAAAAA", // 备注
        checkReason: '同意' // 审核备注

      },

      // 表格信息
      tableData: []
    };
  },
  created () {
    this.tableData = [
      {
        number: 1, // 序号
        name: "设备1", // 设备名称
        num: 1, // 数量
        salePrice: 10, // 销售单价
        saleTotal: 10, // 销售合计
        remark: "啦啦啦" // 备注
      },
      {
        number: 2, // 序号
        name: "设备2", // 设备名称
        num: 2, // 数量
        salePrice: 20, // 销售单价
        saleTotal: 40, // 销售合计
        remark: "啦啦啦" // 备注
      },
      {
        number: 3, // 序号
        name: "设备3", // 设备名称
        num: 3, // 数量
        salePrice: 30, // 销售单价
        saleTotal: 90, // 销售合计
        remark: "啦啦啦" // 备注
      }
    ];
  },
  mounted () { this.ImageToBase64() },
  methods: {
    ImageToBase64 () {
      let image = new Image();
      let imageUrl = 'http://localhost:8080/img/logo.82b9c7a5.jpg'
      image.src = imageUrl
      let that = this;
      image.onload = () => {
        var canvas = document.createElement("canvas")
        canvas.width = image.width;
        canvas.height = image.height;
        var context = canvas.getContext('2d');
        context.drawImage(image, 0, 0, image.width, image.height);
        var quality = 0.8;
        //这里的dataurl就是base64类型
        var dataURL = canvas.toDataURL("image/png", quality);//注意图片格式一定要是png的
        console.log(dataURL);
        this.form.imgUrl = dataURL
      }
    },
    // 
    //这里是官网写的应该是处理图片的代码，照着写就行，没动过
    base64DataURLToArrayBuffer (dataURL) {
      const base64Regex = /^data:image\/(png|jpg|svg|svg\+xml);base64,/;
      if (!base64Regex.test(dataURL)) {
        return false;
      }
      const stringBase64 = dataURL.replace(base64Regex, "");
      let binaryString;
      if (typeof window !== "undefined") {
        binaryString = window.atob(stringBase64);
      } else {
        binaryString = new Buffer(stringBase64, "base64").toString("binary");
      }
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        const ascii = binaryString.charCodeAt(i);
        bytes[i] = ascii;
      }
      return bytes.buffer;
    },
    // 导出word
    exportWord () {
      //这里要引入处理图片的插件，下载docxtemplater后，引入的就在其中了
      var ImageModule = require('docxtemplater-image-module-free');
      var that = this;
      //这里是我的Word路径，在static文件下
      JSZipUtils.getBinaryContent("input.docx", function (error, content) {
        if (error) {
          throw error
        };
        let opts = {}
        opts.centered = true;
        opts.fileType = "docx";
        opts.getImage = (tag) => {
          return that.base64DataURLToArrayBuffer(tag);
        }
        opts.getSize = () => {
          return [600, 400]//这里可更改输出的图片宽和高
        }
        let zip = new PizZip(content);
        let doc = new docxtemplater();
        doc.attachModule(new ImageModule(opts));
        doc.loadZip(zip);
        console.log({});
        doc.setData({
          ...that.form,
          table: that.tableData//我的最外层包裹一切要导出的数据名称
        });
        try {
          doc.render()
        } catch (error) {
          var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
          }
          console.log(JSON.stringify({
            error: e
          }));
          throw error;
        }
        var out = doc.getZip().generate({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })
        saveAs(out, "成绩报表.docx")
      })
    },

  }
};
</script>
<style lang="scss">
* {
  margin: 0;
  padding: 0;
  font-size: 12px;
  font-family: "微软雅黑", "宋体";
  list-style: none;
  box-sizing: border-box;
}
// 清除浮动
.clearfix:after {
  content: "";
  height: 0;
  line-height: 0;
  display: block;
  clear: both;
  visibility: hidden;
}

.clearfix {
  zoom: 1;
}
.quoteContainer {
  .quote_info {
    width: 800px;
    margin: 0 auto;
    .h3_title {
      float: left;
      width: 100%;
      font-size: 20px;
      line-height: 40px;
      text-align: center;
      margin-bottom: 30px;
    }
    .quote_itemBox {
      float: left;
      width: 100%;
      margin-bottom: 20px;
      .quote_item {
        float: left;
        width: 50%;
        display: flex;
        .quote_label {
          width: 100px;
          text-align: right;
          line-height: 32px;
        }
        .quote_p {
          flex: 1;
          line-height: 32px;
        }
      }
    }
  }
  .quote_table {
    padding: 0 20px;
    .table_domQuote {
      .quoteTable {
        font-size: 14px;
        padding-left: 30px;
        border-top: 1px solid #ebeef5;
        .quoteTable_sp1 {
          display: inline-block;
          line-height: 50px;
        }
      }
    }
  }
}
.reasonBox {
  padding: 0 20px 20px 20px;
  .title_checkReason {
    line-height: 50px;
  }
  .checkReasonMain {
    .p_box {
      border: 1px solid #ebeef5;
      padding: 10px 20px;
      .p_checkReason {
        height: 30px;
        line-height: 30px;
      }
    }
  }
}
// 底部按钮
.botmBtnContainer {
  text-align: center;
  padding: 20px;
}
</style>
