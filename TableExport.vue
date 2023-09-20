<template>
  <div class="table_page">
    <el-upload
      :show-file-list="false"
      :before-upload="beforeUpload"
      accept=".xlsx,.xls"
    >
      <el-button type="primary">选择xlxs文件</el-button>
    </el-upload>
    <el-table
      :data="tableData"
      ref="table"
      style="
        width: 50%;
        margin-top: 100px;
        margin-left: 50%;
        transform: translateX(-50%);
      "
      @selection-change="handleSelectionChange"
    >
      <el-table-column type="selection" width="55" />
      <el-table-column property="id" label="ID" />
      <el-table-column property="name" label="姓名" />
      <el-table-column property="sex" label="性别" />
      <el-table-column property="age" label="年龄" />
    </el-table>

    <div
      class="btn_grp"
      style="margin-top: 50px; margin-left: 50%; transform: translateX(-50%)"
    >
      <el-button type="primary" @click="exportSelectData"
        >导出选中的行</el-button
      >
      <el-button type="primary" @click="exportTableData"
        >导出表格数据</el-button
      >
    </div>
  </div>
</template>

<script setup>
import XLSX from "xlsx";
import { ref } from "vue";

const tableData = ref([
  { id: 1, name: "张三", sex: "女", age: 23 },
  { id: 2, name: "小五", sex: "男", age: 13 },
  { id: 3, name: "阿迪", sex: "男", age: 29 },
  { id: 4, name: "闪客", sex: "女", age: 88 },
]);

const selectData = ref([]);
const table = ref();
const beforeUpload = (e) => {
  file2XLSX(e).then((res) => {
    console.log("可以继续对res数据进行二次处理");
    let data = res[0].sheet;
    const table = tableData.value;

    //过滤已存在ID行
    table.forEach((v) => {
      data = data.filter((n) => n.id !== v.id);
    });

    tableData.value.unshift(...data);
  });
  return false;
};

const file2XLSX = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsBinaryString(file);
    reader.onload = function (e) {
      console.log(e, "读取文件成功");
      // 获取读取文件成功的结果值
      const data = e.target.result;
      // XLSX.read解析数据，按照type 的类型解析
      let wb = XLSX.read(data, {
        type: "binary", // 二进制
      });
      console.log(wb, "---->解析后的数据");
      // 存储获取到的数据
      const result = [];
      // 工作表名称的有序列表
      wb.SheetNames.forEach((sheetName) => {
        result.push({
          // 工作表名称
          sheetName: sheetName,
          // 利用 sheet_to_json 方法将 excel 转成 json 数据
          sheet: XLSX.utils.sheet_to_json(wb.Sheets[sheetName]),
        });
      });
      resolve(result);
    };
  });
};

const handleSelectionChange = (e) => {
  selectData.value.push(...e);
};

const exportSelectData = () => {
  // 对表格数据进行整理,添加标题
  let arr = selectData.value.map((item) => {
    return {
      ID: item.id,
      姓名: item.name,
      性别: item.sex,
      年龄: item.age,
    };
  });

  // 将有对象组成的数组转为sheet;
  let sheet = XLSX.utils.json_to_sheet(arr);

  // 新建表格
  let book = XLSX.utils.book_new();

  // 在表格中插入sheet
  XLSX.utils.book_append_sheet(book, sheet, "sheet1");
  console.log(book);
  // 通过xlsx的writeFile方法将文件写入
  return XLSX.writeFile(book, `user${new Date().getTime()}.xls`);
};

const exportTableData = () => {
  let table_copy = table.value.$el;
  // 不是原生dom须在后面加$el   ,,   table dom 转sheet
  let sheet = XLSX.utils.table_to_sheet(table_copy);

  deleteCol(sheet, 0);
  console.log(sheet);
  // 创建新表
  let book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "sheet1");
  console.log(book);
  XLSX.writeFile(book, `user${new Date().getTime()}.xls`);  
};

// 删除指定列
function deleteCol(ws, index) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let col = index; col < range.e.c; col++) {
    for (let row = range.s.r; row <= range.e.r; row++) {
      ws[encodeCell(row, col)] = ws[encodeCell(row, col + 1)];
    }
  }

  range.e.c--;

  ws["!ref"] = XLSX.utils.encode_range(range.s, range.e);
}

function encodeCell(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}
</script>
