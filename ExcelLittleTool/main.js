// 依赖包
var xlsx = require('node-xlsx');
var fs   = require('fs');
var china= require('china-address');

// 常量
var SOURCE_PATH    = '../files/source.xlsx';    // 源文件路径
var TARGET_PATH    = '../files/target.xlsx';    // 生成文件保存路径


/**
 * 主入口函数
 */

function main(address_x, filePath) {
  if (filePath){
    SOURCE_PATH = filePath;
    var splits = filePath.split('/');
    // var splits = filePath.split('\\'); // for windows
    var filename = splits[splits.length-1];
    TARGET_PATH = filePath.replace(filename, 'target.xlsx');
  }
  if (!address_x){
    address_x = 11;
  }else{
    address_x--;
  }

  try{
    // 读取文件
    var obj = xlsx.parse(SOURCE_PATH);

    // 获取第一个表
    var table = obj[0].data;

    // 对收货地址提取和修改
    var parseredData = dataDeal(table, address_x);

    // 对数据进行二进制编码
    var binaryFile = xlsx.build([{name:'target', data: parseredData}]);

    // 保存文件
    fs.writeFile(TARGET_PATH, binaryFile, 'binary', function(error){
      if (error) {
        alert('转换失败，未知错误');
      }
      alert('转换完毕');
    });
  }catch(error){
    alert(error)
  }
}


/**
 * 处理数据表的数据
 * @param data
 * @param address_x
 * @returns {*}
 */
function dataDeal(data, address_x){
  var array = [
    ['省', '市', '区', '收货地址']
  ];
  for (var i = 1; i < data.length; i++) {
    var address = data[i][address_x];
    if (address) {
      var addressDrivion = addressParser(address);
      addressDrivion.push(address);
      array.push(addressDrivion);
    }
  }
  return array;
}

/**
 * 提取地址
 * @param address
 * @returns {*}
 */
function addressParser(address){
  return china.location(address, {postfix: true});
}
