
/*****************************************************************
● gfnExcelExport     : Excel Export 
● gfnExcelImportAll  : Excel Import (헤더 포함)
● gfnExcelImport     : Excel Import (헤더 제외)
******************************************************************/


var pForm = nexacro.Form.prototype;


// Excel Export
/**
 * @class excel export <br>
 * @param {Object} objGrid - Grid Object   
 * @param {String} [sSheetName]   - sheet name
 * @param {String} [sFileName]   - file name
 * @param {strign} default supress
    "suppress" 설정 시 suppress 된 결과대로 1 개 Cell 만 값을 Export 합니다.
   나머지 Row 의 해당 Cell 은 병합되지 않으며 모두 공백으로 처리됩니다.

   "nosuppress" 설정 시 suppress 된 결과를 무시하고 각 Cell 에 모두 값을 Export 합니다.

   "merge" 설정 시 suppress 된 결과대로 1 개 Cell 만 값을 Export 합니다.
   나머지 Row 의 해당 Cell 은 병합되어 처리됩니다.

 * @return N/A
 * @example
 * this.gfnExcelExport(this.grid_export, "SheetName","");
 */
pForm.gfnExcelExport = function (objGrid, sSheetName, sFileName, sSuppress) {
   // Validation
   // Row값이 0일경우 오류처리
   if (objGrid.rowcount == 0) {
      this.gfnAlert("COM_SYS_E0393");  //다운로드할 자료가 없습니다.
      return;
   }

   // Wait Cursor(모래 시계 아이콘) 표시, userwaitcursor 속성값에 관계없이 메서드 실행 
   this.setWaitCursor(true, true);

   try {
      //sSuppress가 null이면 "suppress"
      var pSupppress = sSuppress || "suppress";

      // Excel 저장용 Grid 생성
      // getCurFormatString - false : 변경된 포맷 반영해서 리턴
      var strContents = objGrid.getCurFormatString(false);
      var oGrdExcelExNm = objGrid.name + "_exTemp";

      // 그리드 생성 여부 체크   
      if (!this.isValidObject(oGrdExcelExNm)) {
         var objGrdExcelEx = new Grid();

         objGrdExcelEx.init(oGrdExcelExNm, 0, 0, 0, 0, null, null);
         // Add Object to Parent Form  
         this.addChild(oGrdExcelExNm, objGrdExcelEx);
         objGrdExcelEx.set_autofittype("none");
         objGrdExcelEx.set_autoenter("select");
         objGrdExcelEx.set_selecttype("multirow");
         objGrdExcelEx.set_binddataset(objGrid.binddataset);
         objGrdExcelEx.set_summarytype(objGrid.summarytype);
         objGrdExcelEx.set_formats("<Formats>" + strContents + "</Formats>");

         if (objGrdExcelEx.hasOwnProperty("griduserproperty")) objGrdExcelEx.griduserproperty = objGrid.griduserproperty;

         // Add Object to Parent Form  
         objGrdExcelEx.set_visible(false);
         // Show Object  
         objGrdExcelEx.show();

         for (var k = 3; k >= 0; k--) {
            var sBindTxt = objGrdExcelEx.getCellProperty("boyd", k, "text");
            if (!this.gfnIsNull(sBindTxt)) sBindTxt = sBindTxt.toUpperCase();
            var tmp = objGrdExcelEx.getCellProperty("head", k, "text");
            //개행 -> 띄어쓰기로
            tmp = tmp.replace(/\n/g, ' ');
            if (!this.gfnIsNull(tmp)) tmp = tmp.toUpperCase();

            if (sBindTxt == "BIND:_CHK" || tmp == "ICON" || tmp == "NO") {
               objGrdExcelEx.deleteContentsCol("body", k, false);   // 컬럼 삭제
            }

         }
      }

      // Validation
      var objGrid_excel = this[oGrdExcelExNm];

      var regExp = /[?*:\/\[\]]/g;              //(Excel - 지원하지않는 문자)
      sFileName = this.gfnIsNull(sFileName) ? this.gfnGetDate('time') : sFileName;
      //Excel Sheet Name nullcheck
      sSheetName = this.gfnIsNull(sSheetName) ? "sheet1" : sSheetName;

      sFileName = sFileName.replace(regExp, "");   //파일명 특수문자 제거
      sSheetName = sSheetName.replace(regExp, ""); //시트명 특수문자 제거
      //sheetName 30이상일경우 기본시트명
      if (String(sSheetName).length > 30) {
         sSheetName = "sheet1";
      }

      var svcUrl = "svcUrl::XExportImport";
      this.objExport = null
      this.objExport = new ExcelExportObject();

      this.objExport.objgrid = objGrid_excel;
      this.objExport.set_exporturl(svcUrl);
      this.objExport.addExportItem(nexacro.ExportItemTypes.GRID, objGrid_excel, sSheetName + "!A1", "allband", "allrecord", "nosuppress", "allstyle", "image", "", "both");
      this.objExport.set_exportfilename(sFileName);

      this.objExport.set_exporteventtype("itemrecord");
      this.objExport.set_exportuitype("none");
      this.objExport.set_exporttype(nexacro.ExportTypes.EXCEL2007);
      this.objExport.set_exportmessageprocess("");
      this.objExport.addEventHandler("onsuccess", this.gfnExportOnsuccess, this);
      this.objExport.addEventHandler("onerror", this.gfnExportOnerror, this);

   } catch (err) {
      this.setWaitCursor(false, true);
      trace("★★★excel export error : " + err);
   }
};

// Excel Export
/**
 * @class excel export on sucess <br>
 * @param {Object} obj   
 * @param {Event} e      
 * @return N/A
 * @example
 */
pForm.gfnExportOnsuccess = function () {
   this.setWaitCursor(false, true);
};

/**
 * @class  excel export on error <br>
 * @param {Object} obj   
 * @param {Event} e      
 * @return N/A
 * @example
 */
pForm.gfnExportOnerror = function () {
   this.setWaitCursor(false, true);

   var sId = "EXCEL DOWNLOAD FAIL!!";
   var arrArg = [];
   var sMsgId = sId;
   this.gfnAlert(sId, arrArg, sMsgId, function () { }, "E");
   return;

};

/**
 * @class  excel import( 데이터 헤더포함 ) <br>
 * @param {String} objDs - dataset   
 * @param {String} [sSheet]   - sheet name(default:Sheet1)
 * @param {String} sHead - Head 영역지정   
 * @param {String} [sBody] - body 영역지정(default A2)   
 * @param {String} [sCallback]   - callback 함수
 * @param {String} [sImportId] - import id(callback호출시 필수)   
 * @param {Object} [objForm] - form object(callback호출시 필수)
 * @return N/A
 * @example
 * this.gfnExcelImportAll("dsList","SheetName","A1:G1","A2","fnImportCallback","import",this);
 */
pForm.gfnExcelImportAll = function (objDs, sSheet, sHead, sBody, sCallback, sImportId, objForm) {

   // 세팅 및 Validation
   if (this.gfnIsNull(sSheet)) sSheet = "sheet1";
   if (this.gfnIsNull(sBody)) sBody = "A2";
   if (this.gfnIsNull(sHead)) return false;

   var oDsImport = this.gfnGetDataSet(objDs + "_importEx");   //excel import dataset 생성
   oDsImport.clearData();
   oDsImport.copyData(this[objDs]);

   // 데이터셋에 자동 생성되는 _CHK 컬럼 삭제
   var objColInfo = oDsImport.getColumnInfo("_CHK");
   if (objColInfo != undefined) {
      oDsImport.deleteColumn("_CHK");
   }

   var svcUrl = "svcUrl::XExportImport";
   alert(svcUrl);

   var objImport;

   objImport = new nexacro.ExcelImportObject(objDs + "_ExcelImport", this);
   objImport.set_importurl(svcUrl);
   objImport.set_importtype(nexacro.ImportTypes.EXCEL);
   objImport.outds = oDsImport.name;

   if (!this.gfnIsNull(sCallback)) {
      objImport.callback = sCallback;
      objImport.importid = sImportId;
      objImport.form = objForm;
   }

   var sOutDsName = oDsImport.name + "_outds";

   if (this.isValidObject(sOutDsName)) {
      this.removeChild(sOutDsName);
   }
   var objOutDs = new Dataset();
   objOutDs.name = sOutDsName;
   this.addChild(objOutDs.name, objOutDs);

   objImport.addEventHandler("onsuccess", function (obj, e) {

   }, this);
   objImport.addEventHandler("onerror", function (obj, e) {

   }, this);
   var sParam1 = "[Command=getsheetdata;output=outDs;Head=" + sSheet + "!" + sHead + ";Body=" + sSheet + "!" + sBody + "]";
   var sParam2 = "[" + sOutDsName + "=outDs]";


   objImport.importData("", sParam1, sParam2);
   objImport = null;
};


/**
 * @class  excel import( 데이터 헤더제외 ) <br>
 * @param {String} sDataset - dataset   
 * @param {String} [sSheet]   - sheet name
 * @param {String} [sBody] - body 영역지정   
 * @param {String} [sCallback] - callback 함수   
 * @param {String} [sImportId] - import id(callback호출시 필수)   
 * @param {Object} [objForm] - form object(callback호출시 필수)   
 * @return N/A
 * @example
 * this.gfnExcelImport("dsList","SheetName","A2","fnImportCallback","import",this);
 */
pForm.gfnExcelImport = function (sDataset, sSheet, sBody, sCallback, sImportId, objForm) {

   if (this.gfnIsNull(sSheet)) sSheet = "sheet1";
   if (this.gfnIsNull(sBody)) sBody = "A2";

   var svcUrl = "svcUrl::XExportImport";

   var oDsImport = this.gfnGetDataSet(sDataset + "_importEx");   //excel import dataset 생성
   oDsImport.clearData();
   oDsImport.copyData(this[sDataset]);

   // 데이터셋에 자동 생성되는 _CHK 컬럼 삭제
   var objColInfo = oDsImport.getColumnInfo("_CHK");
   if (objColInfo != undefined) {
      oDsImport.deleteColumn("_CHK");
   }


   var objImport;
   objImport = new nexacro.ExcelImportObject(sDataset + "_ExcelImport", this);
   objImport.set_importurl(svcUrl);
   objImport.set_importtype(nexacro.ImportTypes.EXCEL);
   objImport.outds = oDsImport.name;

   if (!this.gfnIsNull(sCallback)) {
      objImport.callback = sCallback;
      objImport.importid = sImportId;
      objImport.form = objForm;
   }


   //out dataset 생성(차후 onsucess 함수에서 헤더생성하기 위한)
   //var sOutDsName = sDataset+"_outds";   
   var sOutDsName = oDsImport.name + "_outds";

   if (this.isValidObject(sOutDsName)) this.removeChild(sOutDsName);
   var objOutDs = new Dataset();
   objOutDs.name = sOutDsName;
   this.addChild(objOutDs.name, objOutDs);

   objImport.addEventHandler("onsuccess", this.gfnImportOnsuccess, this);
   objImport.addEventHandler("onerror", this.gfnImportAllOnerror, this);

   var sParam = "[command=getsheetdata;output=outDs;body=" + '' + "!" + sBody + ";]";
   var sParam2 = "[" + sOutDsName + "=outDs]";

   objImport.importData("", sParam, sParam2);
   objImport = null;

};

/**
 * @class excel import on success <br>
 * @param {Object} obj   
 * @param {Event} e      
 * @return N/A
 * @example
 */
pForm.gfnImportOnsuccess = function (obj, e) {

   var objImportDs = this.objects[obj.outds];
   var strOrgDs = obj.outds.split("_importEx")[0];
   var objOrgDs = this.objects[strOrgDs];
   var objOutDs = this.objects[obj.outds + "_outds"];

   var sCallback = obj.callback;
   var sImportId = obj.importid;
   var objForm = obj.form;
   var sColumnId;


   //기존 데이터셋의 내용으로 헤더복사
   for (var i = 0; i < objOrgDs.getColCount(); i++) {
      sColumnId = "Column" + i;

      if (sColumnId != objOrgDs.getColID(i)) {
         objOutDs.updateColID(sColumnId, objOrgDs.getColID(i))
      }
   }

   // trace(" objOutDs.saveXML() " + objOutDs.saveXML());

   objOrgDs.clearData();
   objOrgDs.copyData(objOutDs);


   //화면의 callback 함수 호출
   if (!this.gfnIsNull(sCallback)) {

      if (nexacro._isFunction(sCallback)) sCallback.call(this, 0, sImportId, objOrgDs);
      else
         this[sCallback].call(this, 0, sImportId, objOrgDs);
   }
};

/**
 * @class  excel import on error <br>
 * @param {Object} obj   
 * @param {Event} e      
 * @return N/A
 * @example
 */
pForm.gfnImportAllOnerror = function (obj, e) {
   var objOutDs = this.objects[obj.outds + "_outds"];
   var objOrgDs = this.objects[obj.outds];
   var sCallback = obj.callback;
   var sImportId = obj.importid;
   var objForm = obj.form;

   //화면의 callback 함수 호출
   if (!this.gfnIsNull(sCallback)) {
      if (nexacro._isFunction(sCallback)) sCallback.call(this, -1, sImportId);
      else
         this[sCallback].call(this, -1, sImportId);
   }
};