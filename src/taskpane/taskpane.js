/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btn").addEventListener("click",writeData);
  }
});

export async function writeData() {
    // Excel.run((context) => {
    //     const ws= context.workbook.worksheets.getActiveWorksheet();
    //     const range = ws.getRange("A1:A2")
    //     range.values = [[44],[77]]
    //     return context.sync()
    // });
    Excel.run(async (context) => {
        const fileUrl = "https://transfer.sh/get/yszaX2/Mapping_CN-BAR-NBSdb-IP-MTH.xlsx";
        const fileName = "myfile.xlsx";
        const workbook = await context.workbook;
        const sheets = workbook.worksheets;
        // 从 URL 下载文件
        const response = await fetch(fileUrl,{
                                     mode: 'cors',
                                     headers: {'Access-Control-Allow-Origin': "*",
                                                "Access-Control-Allow-Methods": "GET, POST, PUT",
                                                "Access-Control-Allow-Headers": "Content-Type"},
                                    });
        const data = await response.arrayBuffer();
        // 创建一个 Base64 字符串以在 Excel 中加载文件
        const base64Data = btoa(String.fromCharCode.apply(null, new Uint8Array(data)));
        // 将文件加载到 Excel 中
        const worksheet = sheets.getItem(0);
        worksheet.activate();
        worksheet.getRange("A1").select();
        worksheet.paste(base64Data, "xlsx");
        // 重命名工作簿
        workbook.name = fileName;
      await context.sync();
      }).catch((error) => {
        console.error(error);
      });
}
