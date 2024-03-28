<template>
  <div id="app" style="padding: 10px">
    <div style="display: flex; justify-content: center; width: 100%">
      <el-upload
        action
        accept=".xlsx, .xls"
        :auto-upload="false"
        :show-file-list="false"
        :on-change="handle"
      >
        <el-button type="primary" style="margin-right: 10px"
          >导入excel</el-button
        >
      </el-upload>
      <el-button @click="exportWord" type="success">导出word</el-button>
    </div>

    <div v-for="item in resultData" style="margin-top: 30px">
      <h2>{{ item.sheetName }}</h2>
      <el-table
        :data="item.tableData"
        border
        stripe
        highlight-current-row
        style="width: 100%; margin-bottom: 30px"
        max-height="600"
      >
        <el-table-column
          v-for="val in item.tableTitle"
          :prop="val"
          :label="val"
          :key="val"
        >
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { storeToRefs } from "pinia";
import { useNum } from "@/store/first";
import { computed } from "vue";

//导出word 依赖
import docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import JSZipUtils from "jszip-utils";
import { saveAs } from "file-saver";
// 导入excel 依赖
import * as XLSX from "xlsx/xlsx.mjs";
import { ref } from "vue";
// 定义导出word模板变量
// 定义用电量比例
let number1 = null, //全市用电量
  number2 = null,
  number3 = null,
  number4 = null,
  number5 = null,
  number6 = null,
  number7 = null,
  number8 = null,
  number11 = null,
  number12 = null,
  number13 = null,
  number14 = null,
  number15 = null,
  number16 = null,
  number17 = null,
  number18 = null,
  number19 = null,
  number20 = null,
  number21 = null,
  number22 = null;
// 定义一段文字
let dynamicText = undefined;
// 定义商圈 景区
let location1 = undefined,
  location2 = undefined,
  location3 = undefined,
  location4 = undefined,
  location5 = undefined,
  location6 = undefined,
  location7 = undefined,
  location8 = undefined,
  location9 = undefined,
  location10 = undefined,
  location11 = undefined,
  location12 = undefined;
let resultData = ref([]);
// 导入文件
function readFile(file) {
  //文件读取
  return new Promise((resolve) => {
    let reader = new FileReader();
    reader.readAsBinaryString(file); //以二进制的方式读取
    reader.onload = (ev) => {
      resolve(ev.target.result);
    };
  });
}
const handle = async (ev) => {
  let file = ev.raw;
  if (!file) {
    console.log("文件打开失败");
    return;
  } else {
    let tableData = [];
    let tableHeader = [];
    let sheetNames = [];
    let data = await readFile(file);
    let workbook = XLSX.read(data, { type: "binary" }); //解析二进制格式数据
    let worksheet1 = workbook.Sheets[workbook.SheetNames[0]]; //获取第一个Sheet
    let result1 = XLSX.utils.sheet_to_json(worksheet1); //json数据格式

    let worksheet2 = workbook.Sheets[workbook.SheetNames[1]]; //获取第二个Sheet
    let result2 = XLSX.utils.sheet_to_json(worksheet2); //json数据格式

    let worksheet3 = workbook.Sheets[workbook.SheetNames[2]]; //获取第三个Sheet
    let result3 = XLSX.utils.sheet_to_json(worksheet3); //json数据格式

    //处理数据 合计重复数据项
    result1 = [...sumSameRow(result1, "分类电量（万千瓦时）")];
    result2 = [...sumSameRow(result2, "商圈")];
    result3 = [...sumSameRow(result3, "景区名称")];

    firstSheet(result1);
    threeSheet(result3);
    secondSheet(result2);

    // 获取全部数据 用于页面显示
    sheetNames = workbook.SheetNames;
    await Promise.all(
      sheetNames.map(async (item) => {
        let jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[item]);
        tableData.push(jsonData);
        if (jsonData.length > 0) {
          let keys = Object.keys(jsonData[0]);
          tableHeader.push(keys);
        }
      })
    );
    resultData.value = tableData.map((data, index) => {
      return {
        sheetName: sheetNames[index],
        tableData: data,
        tableTitle: tableHeader[index],
      };
    });
  }
};

// 计算第一个sheet数据
function firstSheet(data) {
  countAllCity(data);
  countOneIndustry(data);
  countTwoIndustry(data);
  countThreeIndustry(data);
  countAllPeople(data);
  countAllIndustry(data);
}

// 计算第三个sheet数据 景区
function threeSheet(data) {
  findTopToIndustry(data);
}

// 计算第二个sheet数据 商圈
function secondSheet(data) {
  findTopToShop(data);
}

// 计算全市用电量数据
function countAllCity(data) {
  let { firstallEle, secondallEle } = sumSameYearData(data, "全社会用电总计");
  // 给导出数据赋值（全市用电量） 将千瓦转换成亿千瓦 保留小数点两位
  number1 = (secondallEle / 100000000).toFixed(2);
  // 给导出数据赋值（全市用电量）同比 保留小数点两位
  if (secondallEle / firstallEle - 1 > 0) {
    number2 = "增长" + ((secondallEle / firstallEle - 1) * 100).toFixed(2);
  } else {
    number2 =
      "下降" + Math.abs(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  }
}

// 计算第一产业用电量数据
function countOneIndustry(data) {
  let { firstallEle, secondallEle } = sumSameYearData(data, "第一产业");
  // 给导出数据赋值（第一产业）同比 保留小数点两位
  if (secondallEle / firstallEle - 1 > 0) {
    number3 = "增长" + ((secondallEle / firstallEle - 1) * 100).toFixed(2);
  } else {
    number3 =
      "下降" + Math.abs(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  }
}

// 计算第二产业用电量数据
function countTwoIndustry(data) {
  let { firstallEle, secondallEle } = sumSameYearData(data, "第二产业");
  // 给导出数据赋值（第二产业）同比 保留小数点两位
  if (secondallEle / firstallEle - 1 > 0) {
    number4 = "增长" + ((secondallEle / firstallEle - 1) * 100).toFixed(2);
  } else {
    number4 =
      "下降" + Math.abs(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  }
}

// 计算第三产业用电量数据
function countThreeIndustry(data) {
  let { firstallEle, secondallEle } = sumSameYearData(data, "第三产业");
  // 给导出数据赋值（第三产业）同比 保留小数点两位
  if (secondallEle / firstallEle - 1 > 0) {
    number5 = "增长" + ((secondallEle / firstallEle - 1) * 100).toFixed(2);
  } else {
    number5 =
      "下降" + Math.abs(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  }
}

// 计算居民用电量数据
function countAllPeople(data) {
  let { firstallEle, secondallEle } = sumSameYearData(
    data,
    "B、城乡居民生活用电合计"
  );
  // 给导出数据赋值（居民用电量）同比 保留小数点两位
  if (secondallEle / firstallEle - 1 > 0) {
    number6 = "增长" + ((secondallEle / firstallEle - 1) * 100).toFixed(2);
  } else {
    number6 =
      "下降" + Math.abs(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  }
}

// 计算旅游相关行业用电量数据
function countAllIndustry(data) {
  let trafficForm = sumSameYearData(data, "四、交通运输、仓储和邮政业");
  let retailForm = sumSameYearData(data, "六、批发和零售业");
  let stayForm = sumSameYearData(data, "七、住宿和餐饮业");
  let serviceForm = sumSameYearData(data, "十、租赁和商务服务业");
  //旅游行业 往年数据之和
  let industryPre =
    trafficForm.firstallEle +
    retailForm.firstallEle +
    stayForm.firstallEle +
    serviceForm.firstallEle;
  //旅游行业 今年数据之和
  let industrycur =
    trafficForm.secondallEle +
    retailForm.secondallEle +
    stayForm.secondallEle +
    serviceForm.secondallEle;
  // 交通同比
  let trafficRatio = (
    (trafficForm.secondallEle / trafficForm.firstallEle - 1) *
    100
  ).toFixed(2);
  // 批发零售同比
  let retailRatio = (
    (retailForm.secondallEle / retailForm.firstallEle - 1) *
    100
  ).toFixed(2);
  // 住宿餐饮同比
  let stayRatio = (
    (stayForm.secondallEle / stayForm.firstallEle - 1) *
    100
  ).toFixed(2);
  // 租赁服务同比
  let serviceRatio = (
    (serviceForm.secondallEle / serviceForm.firstallEle - 1) *
    100
  ).toFixed(2);

  // 给导出数据赋值（旅游相关行业用电量） 以上四个行业之和 将千瓦转换成亿千瓦 保留小数点两位
  number7 = (industrycur / 100000000).toFixed(2);
  // 给导出数据赋值（旅游相关行业用电量）同比 保留小数点两位
  if (industrycur / industryPre - 1 > 0) {
    number8 = "增长" + ((industrycur / industryPre - 1) * 100).toFixed(2);
  } else {
    number8 =
      "下降" + Math.abs(((industrycur / industryPre - 1) * 100).toFixed(2));
  }
  // 定义旅游行业同比数组
  let ratioForm = [trafficRatio, retailRatio, stayRatio, serviceRatio];
  // 插入文字的模板字符串数组
  let textForm = [
    `交通运输、仓储和邮政业（${trafficRatio}%）、`,
    `批发和零售业（${retailRatio}%）、`,
    `住宿和餐饮业（${stayRatio}%）、`,
    `租赁和商务服务业（${serviceRatio}%）、`,
  ];
  // 同比增长的index数组
  let ratioIndex = [];
  // 同比增长行业及同比比例字符串
  let textIndustry = "";
  for (let i = 0; i < ratioForm.length; i++) {
    if (ratioForm[i] > 0) {
      ratioIndex.push(i);
    }
  }
  ratioIndex.map((v) => {
    textIndustry += textForm[v];
  });
  // 有增长行业显示增长行业 没有增长行业显示空白
  if (ratioIndex.length != 0) {
    // 去掉最后一个顿号、
    textIndustry = textIndustry.substring(0, textIndustry.lastIndexOf("、"));
    dynamicText = `，其中${textIndustry}用电量均呈同比增长态势`;
  } else {
    dynamicText = textIndustry;
  }
}

//计算相同年份的用电量数据之和 公用
function sumSameYearData(data, type) {
  // 计算的这一行的数据
  let totalEle = null;
  // 找到目标这一行数据
  for (let i = 0; i < data.length; i++) {
    if (
      Object.values(data[i]).some((item) => {
        return item == type;
      })
    ) {
      totalEle = data[i];
    }
  }
  //相同年份数据的和
  let allEle = 0;
  //第一组相同的和
  let firstallEle = 0;
  //第二组相同的和
  let secondallEle = 0;
  for (let i = 0; i < Object.keys(totalEle).length; i++) {
    if (
      (Object.keys(totalEle)[i + 1] &&
        Object.keys(totalEle)[i].substring(0, 4) ==
          Object.keys(totalEle)[i + 1].substring(0, 4)) ||
      (Object.keys(totalEle)[i - 1] &&
        Object.keys(totalEle)[i].substring(0, 4) ==
          Object.keys(totalEle)[i - 1].substring(0, 4))
    ) {
      allEle += Object.values(totalEle)[i];
      //判断第一组相同的年份
      if (
        Object.keys(totalEle)[i + 1] &&
        Object.keys(totalEle)[i].substring(0, 4) !=
          Object.keys(totalEle)[i + 1].substring(0, 4) &&
        Object.keys(totalEle)[i].substring(0, 4) ==
          Object.keys(totalEle)[i - 1].substring(0, 4)
      ) {
        firstallEle = allEle;
        allEle = 0;
      }
      // 判断第二组相同的年份
      if (!Object.keys(totalEle)[i + 1]) {
        secondallEle = allEle;
      }
    }
  }
  return { firstallEle, secondallEle };
}

//计算出景区的top3
function findTopToIndustry(data) {
  let { topCurEleName, topCurEle, topRatioEleName, topRatioEle } =
    sumEveryRowSameYearData(data, "景区名称");
  //用电量名称前三
  location1 = topCurEleName[0];
  location2 = topCurEleName[1];
  location3 = topCurEleName[2];
  // 单位  万千瓦
  number11 = (topCurEle[0] / 10000).toFixed(2);
  number12 = (topCurEle[1] / 10000).toFixed(2);
  number13 = (topCurEle[2] / 10000).toFixed(2);

  // 同比前三
  location4 = topRatioEleName[0];
  location5 = topRatioEleName[1];
  location6 = topRatioEleName[2];
  // 单位  百分比
  number14 = topRatioEle[0];
  number15 = topRatioEle[1];
  number16 = topRatioEle[2];
}

//计算出商圈的top3
function findTopToShop(data) {
  let { topCurEleName, topCurEle, topRatioEleName, topRatioEle } =
    sumEveryRowSameYearData(data, "商圈");
  //用电量名称前三
  location7 = topCurEleName[0];
  location8 = topCurEleName[1];
  location9 = topCurEleName[2];
  // 单位  万千瓦
  number17 = (topCurEle[0] / 10000).toFixed(2);
  number18 = (topCurEle[1] / 10000).toFixed(2);
  number19 = (topCurEle[2] / 10000).toFixed(2);

  // 同比前三
  location10 = topRatioEleName[0];
  location11 = topRatioEleName[1];
  location12 = topRatioEleName[2];
  // 单位  百分比
  number20 = topRatioEle[0];
  number21 = topRatioEle[1];
  number22 = topRatioEle[2];
}

// 计算每一行数据的相同年份的和 公用
function sumEveryRowSameYearData(data, type) {
  // 每一行去年用电量之和
  let preYearSum = [];
  // 每一行今年用电量之和
  let curYearSum = [];
  // 每一行用电量同比
  let ratioBoth = [];
  // 每一行商圈或景区名字
  let nameForm = [];
  data.map((v) => {
    //相同年份数据的和
    let allEle = 0;
    //第一组相同的和
    let firstallEle = 0;
    //第二组相同的和
    let secondallEle = 0;
    // 商圈或景区名字
    nameForm.push(v[type]);
    for (let i = 0; i < Object.keys(v).length; i++) {
      if (
        (Object.keys(v)[i + 1] &&
          Object.keys(v)[i].substring(0, 4) ==
            Object.keys(v)[i + 1].substring(0, 4)) ||
        (Object.keys(v)[i - 1] &&
          Object.keys(v)[i].substring(0, 4) ==
            Object.keys(v)[i - 1].substring(0, 4))
      ) {
        allEle += Object.values(v)[i];
        //判断第一组相同的年份
        if (
          Object.keys(v)[i + 1] &&
          Object.keys(v)[i].substring(0, 4) !=
            Object.keys(v)[i + 1].substring(0, 4) &&
          Object.keys(v)[i].substring(0, 4) ==
            Object.keys(v)[i - 1].substring(0, 4)
        ) {
          firstallEle = allEle;
          preYearSum.push(firstallEle);
          allEle = 0;
        }
        // 判断第二组相同的年份
        if (!Object.keys(v)[i + 1]) {
          secondallEle = allEle;
          curYearSum.push(secondallEle);
        }
      }
    }
    // 同比
    ratioBoth.push(((secondallEle / firstallEle - 1) * 100).toFixed(2));
  });
  // 找到最大的用电量前三{topThreeData,arrayIndex}
  let curForm = sortTopToData(curYearSum);
  //前三用电量
  let topCurEle = curForm.topThreeData;
  //前三用电量名字
  let topCurEleName = [];
  curForm.arrayIndex.map((v) => {
    topCurEleName.push(nameForm[v]);
  });
  // 找到最大的同比前三
  let ratioForm = sortTopToData(ratioBoth.map(Number));
  //前三同比
  let topRatioEle = ratioForm.topThreeData;
  //前三同比名字
  let topRatioEleName = [];
  ratioForm.arrayIndex.map((v) => {
    topRatioEleName.push(nameForm[v]);
  });
  return { topCurEleName, topCurEle, topRatioEleName, topRatioEle };
}

// 排序 找出今年同比最高的前三  倒序 公用
function sortTopToData(data) {
  let oldArray = [...data];
  let array = [];
  oldArray.map((item) => {
    if (item === item && item != "Infinity") {
      array.push(item);
    }
  });
  for (let i = 0; i < array.length - 1; i++) {
    for (let j = 0; j < array.length - i - 1; j++) {
      // 1.对每一个值和它的下一个值进行比较
      if (array[j] < array[j + 1]) {
        // 如果第一个值更多，则将其赋予自定义计数值 count
        let count = array[j];
        // 反复交换
        array[j] = array[j + 1];
        array[j + 1] = count;
      }
    }
  }
  // 最大的三个值
  let topThreeData = array.slice(0, 3);
  // 最大值的索引
  let arrayIndex = [];
  //依次找到最大值在原数组中的位置
  topThreeData.map((v) => {
    arrayIndex.push(data.indexOf(v));
  });
  return { topThreeData, arrayIndex };
}

// 计算sheet中相同数据行合计
function sumSameRow(data, type) {
  // 创建一个对象来保存累加值
  let result = {};
  // 遍历数据对象，并累加相同景区名称的值
  data.forEach(function (obj) {
    let newName = obj[type];
    let valueObj = {};

    // 累加当前对象的每个日期数据
    for (let key in obj) {
      if (key !== type && key.indexOf("20") !== -1) {
        valueObj[key] = obj[key];
      }
    }

    if (!result[newName]) {
      result[newName] = valueObj;
    } else {
      // 累加当前对象的每个日期数据到result对象中对应名称的属性的相应日期数据上
      for (let dateKey in valueObj) {
        result[newName][dateKey] += valueObj[dateKey];
      }
    }
  });

  // 生成新的数组对象，其中每个对象只包含具有唯一名称的行
  let newArray = [];
  for (let name in result) {
    let newObj = {};
    newObj[type] = name;
    newObj = Object.assign(newObj, result[name]);
    newArray.push(newObj);
  }
  // 输出新的数组对象
  return newArray;
}

//导出word
function exportWord(index, row) {
  // 读取并获得模板文件的二进制内容
  JSZipUtils.getBinaryContent("/outputFile.docx", function (error, content) {
    // 抛出异常
    if (error) {
      throw error;
    }

    // 创建一个PizZip实例，内容为模板的内容
    let zip = new PizZip(content);
    // 创建并加载docxtemplater实例对象
    let doc = new docxtemplater().loadZip(zip);
    // 设置模板变量的值，对象的键需要和模板上的变量名一致，值就是你要放在模板上的值
    let docxData = {
      number1: number1,
      number2: number2,
      number3: number3,
      number4: number4,
      number5: number5,
      number6: number6,
      number7: number7,
      number8: number8,
      number11: number11,
      number12: number12,
      number13: number13,
      number14: number14,
      number15: number15,
      number16: number16,
      number17: number17,
      number18: number18,
      number19: number19,
      number20: number20,
      number21: number21,
      number22: number22,
      dynamicText: dynamicText,
      location1: location1,
      location2: location2,
      location3: location3,
      location4: location4,
      location5: location5,
      location6: location6,
      location7: location7,
      location8: location8,
      location9: location9,
      location10: location10,
      location11: location11,
      location12: location12,
    };
    doc.setData({
      ...docxData,
    });

    try {
      // 用模板变量的值替换所有模板变量
      doc.render();
    } catch (error) {
      // 抛出异常
      let e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties,
      };
      console.log(JSON.stringify({ error: e }));
      throw error;
    }

    // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
    let out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    // 将目标文件对象保存为目标类型的文件，并命名
    saveAs(out, `用电信息报告.docx`);
  });
}

// 方法一
//实例化仓库
// const num = useNum();
// console.log(num.age);
// // 方法二
// //storeToRefs 解构
// const { name, age } = storeToRefs(num);
// console.log(age);
// // 方法三
// // computed 解构
// const newName = computed(() => num.name);
// const newAge = computed(() => num.age);
// console.log(newName);
</script>

<style lang="scss" scoped>
</style>