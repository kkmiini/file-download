////exceljs library 使用
let mergedMap = new Map();
let fileNames = [];


document.getElementById("files").addEventListener("change", function (e) {
  const files = e.target.files;


  mergedMap.clear();
  fileNames = [];


  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    fileNames.push(file.name);


    if (file) {
      const reader = new FileReader();
      reader.onload = function (event) {
        const lines = event.target.result.split("\n");


        lines.forEach((line) => {
          if (line.includes("tbboot")) return;
          if (line.includes("_")) return;


          //const versionIndex = line.indexOf(line.match(/[\s]/) + "version");
          const versionkeyword = " version";
          const versionIndex = line.indexOf(versionkeyword);
          //const words = line.split(" ");
          //const match = line.indexOf(line.match(/[\s]/));
          //const newmatch = line.indexOf(line.match(/[\s]/), match +1);
          //const api = line.substring(0, match);
          //const version = line.substring(newmatch, line.indexOf(','));


          //const api = words[0];
          //const start = line.indexOf(words[2]);
          //const version = line.substring(start, line.indexOf(',') ).trim();


          if (versionIndex !== -1) {
            const api = line.substring(0, versionIndex).trim();
            const rest = line.substring(versionIndex + versionkeyword.length).trim();


            const versionMatch = rest.match(/\d+(\.\d+){1,2}\(\d+\)/);
            if (!versionMatch) return;
            const version = versionMatch[0];
            //const rest = sline.substring(versionIndex + 8).trim();
            //const version = rest.substring(0, rest.indexOf(',')).trim();
            /*const api = line.match(/^\s*(\s+)/);
                        if(api) {
                            mergedMap.set("api", api[1]);
                        }
                        const version = line.match(/^\s*(\s+)/);
                        if(version) {
                            mergedMap.set("version", version[0]);
                        } */
            if (api.includes(" no")) return;
            /*if (api.indexOf("version") != -1) {
                            versionIndex = line.indexOf(newmatch);
                            rest = line.substring(versionIndex + 1).trim();
                            version = rest.substring(0, rest.indexOf(',')).trim();
                        } */

            if (!mergedMap.has(api)) {
              mergedMap.set(api, Array(file.length).fill(""));
            }
            const versionArray = mergedMap.get(api);
            versionArray[i] = version;
          }
        });
      };
      reader.readAsText(file);
    }
  }
});

function compareVersions(version1, version2) {
    const parseVersion = (version) => {
      const match = version.match(/\d+(\.\d+){1,2}\(\d+\)/);
      if (!match) return null;
  
      const [main, patch] = match[0].split('(');
      const mainParts = main.split('.').map(Number);
      const patchNumber = parseInt(patch.replace(')', ''));
  
      return { mainParts, patchNumber };
    };
  
    const v1 = parseVersion(version1);
    const v2 = parseVersion(version2);
  
    if (!v1 || !v2) return 0;
  
    
    for (let i = 0; i < Math.max(v1.mainParts.length, v2.mainParts.length); i++) {
      const num1 = v1.mainParts[i] || 0;
      const num2 = v2.mainParts[i] || 0;
      if (num1 > num2) return 1;
      if (num1 < num2) return -1;
    }
  
    if (v1.patchNumber > v2.patchNumber) return 1;
    if (v1.patchNumber < v2.patchNumber) return -1;
    return 0;
}
  
  async function download() {
    const resultArray = [];
  
    for (const [api, version] of mergedMap.entries()) {
      const row = { API: api };
      fileNames.forEach((name, vi) => {
        row[name] = version[vi];
      });
  
      const values = fileNames.map((name, vi) => version[vi] || "");
  
      const isAllSame = values.every((value) => value === values[0]);
  
      row["y/n"] = isAllSame ? "y" : "n";
  
      
      let maxVersion = values[0];
      let maxFileName = fileNames[0];
      
      values.forEach((value, index) => {
        if (compareVersions(value, maxVersion) > 0) {
          maxVersion = value;
          maxFileName = fileNames[index];
        }
      }); 

      const isAllEqual = values.every((value) => compareVersions(value, maxVersion) === 0);
      row["new"] = isAllEqual ? "" : maxFileName;
  
      //row["new"] = maxFileName;
  
      resultArray.push(row);
    }
  
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("result");
  
    const columns = [{ header: "API", key: "API", width: 15 }];
    fileNames.forEach((name) => {
      columns.push({ header: name, key: name, width: 10 });
    });
    columns.push({ header: "NEW", key: "new", width: 40 });
    columns.push({ header: "y/n", key: "y/n", width: 5 });
  
    worksheet.columns = columns;
  
    resultArray.forEach((row) => {
      const newRow = worksheet.addRow(row);
  
      if (row["y/n"] === "n") {
        newRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFFF00'
          }
        };
      }
    });
  
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `패치 목록_SUMINOE_${new Date().toLocaleDateString("ja-JP")}.xlsx`;
    link.click();
  }
  