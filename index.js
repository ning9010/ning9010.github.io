function refreshPage() {
  location.reload();
}

var importButton = document.getElementById("input-mx");

importButton.addEventListener("click", function () {
  var input = document.createElement("input");
  input.type = "file";
  input.accept = '.xls, .xlsx';

  input.onchange = function (event) {
    var file = event.target.files[0]
    var reader = new FileReader()
    var array = ['宝安北', '宝安南', '宝安中', '光明区', '龙华区', '南山区', '福田', '罗盐', '龙岗东', '龙岗西', '深汕']
    let array1 = []
    let array2 = []
    let array3 = []
    let array4 = []
    reader.onload = function (e) {
      var data = e.target.result;
      var wb = XLSX.read(data, { type: "binary" });
      var jsonSheetName = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      for (let i = 0; i < jsonSheetName.length; i++) {
        if (
          jsonSheetName[i].维护代维 === "润建" &&
          jsonSheetName[i].通报类型.includes("e") &&
          !jsonSheetName[i].通报类型.includes("点")
        ) {
          array1.push(jsonSheetName[i].联系号码 + '-' + jsonSheetName[i].网优分区 + '-' + jsonSheetName[i].投诉时间)
        }
        if (
          jsonSheetName[i].维护代维 === "宜通" &&
          jsonSheetName[i].通报类型.includes("e") &&
          !jsonSheetName[i].通报类型.includes("点")
        ) {
          array2.push(jsonSheetName[i].联系号码 + '-' + jsonSheetName[i].网优分区 + '-' + jsonSheetName[i].投诉时间)
        }
        if (
          jsonSheetName[i].维护代维 === "长讯" &&
          jsonSheetName[i].通报类型.includes("e") &&
          !jsonSheetName[i].通报类型.includes("点")
        ) {
          array3.push(jsonSheetName[i].联系号码 + '-' + jsonSheetName[i].网优分区 + '-' + jsonSheetName[i].投诉时间)
        }
        if (
          jsonSheetName[i].维护代维 === "怡创" &&
          jsonSheetName[i].通报类型.includes("e") &&
          !jsonSheetName[i].通报类型.includes("点")
        ) {
          array4.push(jsonSheetName[i].联系号码 + '-' + jsonSheetName[i].网优分区 + '-' + jsonSheetName[i].投诉时间)
        }
      }

      let obj1 = {}
      let obj2 = {}
      let obj3 = {}
      let obj4 = {}

      for (let i = 0; i < array1.length; i++) {
        let count = 0;
        for (let j = 0; j < array1.length; j++) {
          if (array1[j].split("-")[0] == array1[i].split("-")[0] && j != i) {
            count++
          }
        }
        if (count != 0 && array1[i].split("-")[1] === '宝安北') {
          obj1[array1[i].split("-")[0]] = count
        } else if (count != 0 && array1[i].split("-")[1] === '宝安南') {
          obj2[array1[i].split("-")[0]] = count
        }
        else if (count != 0 && array1[i].split("-")[1] === '宝安中') {
          obj3[array1[i].split("-")[0]] = count
        }
        else if (count != 0 && array1[i].split("-")[1] === '光明区') {
          obj4[array1[i].split("-")[0]] = count
        }
      }

      let obj5 = {}
      let obj6 = {}

      for (let i = 0; i < array2.length; i++) {
        let count = 0;
        for (let j = 0; j < array2.length; j++) {
          if (array2[j].split("-")[0] == array2[i].split("-")[0] && j != i) {
            count++
          }
        }
        if (count != 0 && array2[i].split("-")[1] === '龙华区') {
          obj5[array2[i].split("-")[0]] = count
        } else if (count != 0 && array2[i].split("-")[1] === '南山区') {
          obj6[array2[i].split("-")[0]] = count
        }
      }

      let obj7 = {}
      let obj8 = {}

      for (let i = 0; i < array3.length; i++) {
        let count = 0;
        for (let j = 0; j < array3.length; j++) {
          if (array3[j].split("-")[0] == array3[i].split("-")[0] && j != i) {
            count++
          }
        }
        if (count != 0 && array3[i].split("-")[1] === '福田') {
          obj7[array3[i].split("-")[0]] = count
        } else if (count != 0 && array3[i].split("-")[1] === '罗盐') {
          obj8[array3[i].split("-")[0]] = count
        }
      }

      let obj9 = {}
      let obj10 = {}
      let obj11 = {}

      for (let i = 0; i < array4.length; i++) {
        let count = 0;
        for (let j = 0; j < array4.length; j++) {
          if (array4[j].split("-")[0] == array4[i].split("-")[0] && j != i) {
            count++
          }
        }
        if (count != 0 && array4[i].split("-")[1] === '龙岗东') {
          obj9[array4[i].split("-")[0]] = count
        } else if (count != 0 && array4[i].split("-")[1] === '龙岗西') {
          obj10[array4[i].split("-")[0]] = count
        } else if (count != 0 && array4[i].split("-")[1] === '深汕') {
          obj11[array4[i].split("-")[0]] = count
        }
      }

      let ct1 = 0
      let ct2 = 0
      let ct3 = 0
      let ct4 = 0
      for (i = 0; i < Object.keys(obj1).length; i++) {
        ct1 = ct1 + Object.values(obj1)[i]
      }
      for (i = 0; i < Object.keys(obj2).length; i++) {
        ct2 = ct2 + Object.values(obj2)[i]
      }
      for (i = 0; i < Object.keys(obj3).length; i++) {
        ct3 = ct3 + Object.values(obj3)[i]
      }
      for (i = 0; i < Object.keys(obj4).length; i++) {
        ct4 = ct4 + Object.values(obj4)[i]
      }

      let ct5 = 0
      let ct6 = 0

      for (i = 0; i < Object.keys(obj5).length; i++) {
        ct5 = ct5 + Object.values(obj5)[i]
      }
      for (i = 0; i < Object.keys(obj6).length; i++) {
        ct6 = ct6 + Object.values(obj6)[i]
      }

      let ct7 = 0
      let ct8 = 0

      for (i = 0; i < Object.keys(obj7).length; i++) {
        ct7 = ct7 + Object.values(obj7)[i]
      }
      for (i = 0; i < Object.keys(obj8).length; i++) {
        ct8 = ct8 + Object.values(obj8)[i]
      }

      let ct9 = 0
      let ct10 = 0
      let ct11 = 0
      for (i = 0; i < Object.keys(obj9).length; i++) {
        ct9 = ct9 + Object.values(obj9)[i]
      }
      for (i = 0; i < Object.keys(obj10).length; i++) {
        ct10 = ct10 + Object.values(obj10)[i]
      }
      for (i = 0; i < Object.keys(obj11).length; i++) {
        ct11 = ct11 + Object.values(obj11)[i]
      }

      document.getElementById("chongtou-0").textContent = ct1
      document.getElementById("chongtou-1").textContent = ct2
      document.getElementById("chongtou-2").textContent = ct3
      document.getElementById("chongtou-3").textContent = ct4

      document.getElementById("chongtou-4").textContent = ct5
      document.getElementById("chongtou-5").textContent = ct6

      document.getElementById("chongtou-6").textContent = ct7
      document.getElementById("chongtou-7").textContent = ct8

      document.getElementById("chongtou-8").textContent = ct9
      document.getElementById("chongtou-9").textContent = ct10
      document.getElementById("chongtou-10").textContent = ct11



      let count_zbei = 0
      let count_znan = 0
      let count_zong = 0
      let count_zgua = 0

      for (let i = 0; i < array1.length; i++) {
        if (array1[i].split("-")[1] === '宝安北') {
          count_zbei++
        } else if (array1[i].split("-")[1] === '宝安南') {
          count_znan++
        } else if (array1[i].split("-")[1] === '宝安中') {
          count_zong++
        } else if (array1[i].split("-")[1] === '光明区') {
          count_zgua++
        }
      }
      document.getElementById("zongliang-0").textContent = count_zbei
      document.getElementById("zongliang-1").textContent = count_znan
      document.getElementById("zongliang-2").textContent = count_zong
      document.getElementById("zongliang-3").textContent = count_zgua

      let count_zlong = 0
      let count_znans = 0

      for (let i = 0; i < array2.length; i++) {
        if (array2[i].split("-")[1] === '龙华区') {
          count_zlong++
        } else if (array2[i].split("-")[1] === '南山区') {
          count_znans++
        }
      }

      document.getElementById("zongliang-4").textContent = count_zlong
      document.getElementById("zongliang-5").textContent = count_znans

      let count_zfutian = 0
      let count_zluoyan = 0

      for (let i = 0; i < array3.length; i++) {
        if (array3[i].split("-")[1] === '福田') {
          count_zfutian++
        } else if (array3[i].split("-")[1] === '罗盐') {
          count_zluoyan++
        }
      }

      document.getElementById("zongliang-6").textContent = count_zfutian
      document.getElementById("zongliang-7").textContent = count_zluoyan

      let count_zdong = 0
      let count_zxi = 0
      let count_zshan = 0

      for (let i = 0; i < array4.length; i++) {
        if (array4[i].split("-")[1] === '龙岗东') {
          count_zdong++
        } else if (array4[i].split("-")[1] === '龙岗西') {
          count_zxi++
        } else if (array4[i].split("-")[1] === '深汕') {
          count_zshan++
        }
      }

      document.getElementById("zongliang-8").textContent = count_zdong
      document.getElementById("zongliang-9").textContent = count_zxi
      document.getElementById("zongliang-10").textContent = count_zshan

      if (count_zbei != 0) {
        document.getElementById("chongfulv-0").textContent = Number(ct1 / count_zbei * 100).toFixed(2) + '%'
      }
      if (count_znan != 0) {
        document.getElementById("chongfulv-1").textContent = Number(ct2 / count_znan * 100).toFixed(2) + '%'
      }
      if (count_zong != 0) {
        document.getElementById("chongfulv-2").textContent = Number(ct3 / count_zong * 100).toFixed(2) + '%'
      }
      if (count_zgua != 0) {
        document.getElementById("chongfulv-3").textContent = Number(ct4 / count_zgua * 100).toFixed(2) + '%'
      }


      if (count_zlong != 0) {
        document.getElementById("chongfulv-4").textContent = Number(ct5 / count_zlong * 100).toFixed(2) + '%'
      }
      if (count_znans != 0) {
        document.getElementById("chongfulv-5").textContent = Number(ct6 / count_znans * 100).toFixed(2) + '%'
      }

      if (count_zfutian != 0) {
        document.getElementById("chongfulv-6").textContent = Number(ct7 / count_zfutian * 100).toFixed(2) + '%'
      }
      if (count_zluoyan != 0) {
        document.getElementById("chongfulv-7").textContent = Number(ct8 / count_zluoyan * 100).toFixed(2) + '%'
      }

      if (count_zdong != 0) {
        document.getElementById("chongfulv-8").textContent = Number(ct9 / count_zdong * 100).toFixed(2) + '%'
      }
      if (count_zxi != 0) {
        document.getElementById("chongfulv-9").textContent = Number(ct10 / count_zxi * 100).toFixed(2) + '%'
      }
      if (count_zshan != 0) {
        document.getElementById("chongfulv-10").textContent = Number(ct11 / count_zshan * 100).toFixed(2) + '%'
      }


      document.getElementById("zongliang-rj").textContent = parseInt(document.getElementById("zongliang-0").textContent) + parseInt(document.getElementById("zongliang-1").textContent) + parseInt(document.getElementById("zongliang-2").textContent) + parseInt(document.getElementById("zongliang-3").textContent);
      document.getElementById("zongliang-yc").textContent = parseInt(document.getElementById("zongliang-4").textContent) + parseInt(document.getElementById("zongliang-5").textContent);
      document.getElementById("zongliang-hs").textContent = parseInt(document.getElementById("zongliang-6").textContent) + parseInt(document.getElementById("zongliang-7").textContent);
      document.getElementById("zongliang-hx").textContent = parseInt(document.getElementById("zongliang-8").textContent) + parseInt(document.getElementById("zongliang-9").textContent) + parseInt(document.getElementById("zongliang-10").textContent);

      document.getElementById("chongtou-rj").textContent = parseInt(document.getElementById("chongtou-0").textContent) + parseInt(document.getElementById("chongtou-1").textContent) + parseInt(document.getElementById("chongtou-2").textContent) + parseInt(document.getElementById("chongtou-3").textContent);
      document.getElementById("chongtou-yc").textContent = parseInt(document.getElementById("chongtou-4").textContent) + parseInt(document.getElementById("chongtou-5").textContent);
      document.getElementById("chongtou-hs").textContent = parseInt(document.getElementById("chongtou-6").textContent) + parseInt(document.getElementById("chongtou-7").textContent);
      document.getElementById("chongtou-hx").textContent = parseInt(document.getElementById("chongtou-8").textContent) + parseInt(document.getElementById("chongtou-9").textContent) + parseInt(document.getElementById("chongtou-10").textContent);

      if (document.getElementById("zongliang-rj").textContent != 0) {
        document.getElementById("chongfulv-rj").textContent = Number(parseInt(document.getElementById("chongtou-rj").textContent) / parseInt(document.getElementById("zongliang-rj").textContent) * 100).toFixed(2) + '%'
      }
      if (document.getElementById("zongliang-yc").textContent != 0) {
        document.getElementById("chongfulv-yc").textContent = Number(parseInt(document.getElementById("chongtou-yc").textContent) / parseInt(document.getElementById("zongliang-yc").textContent) * 100).toFixed(2) + '%'
      }
      if (document.getElementById("zongliang-hs").textContent != 0) {
        document.getElementById("chongfulv-hs").textContent = Number(parseInt(document.getElementById("chongtou-hs").textContent) / parseInt(document.getElementById("zongliang-hs").textContent) * 100).toFixed(2) + '%'
      }
      if (document.getElementById("zongliang-hx").textContent != 0) {
        document.getElementById("chongfulv-hx").textContent = Number(parseInt(document.getElementById("chongtou-hx").textContent) / parseInt(document.getElementById("zongliang-hx").textContent) * 100).toFixed(2) + '%'
      }
    }
    reader.readAsBinaryString(file);
  }
  input.click();
})