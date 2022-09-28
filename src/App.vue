<template>
  
  <div class="container-fluid">
    <div class="row">
      <div class="col">
        <div class="card">
        <div class="card-header d-flex justify-content-between align-items-center">
          <b>編輯檔案清單</b>
          <div v-if="currentEditFileRow === -1">
            <button type="button" style="margin-right:5px;" class="btn btn-success btn-sm" v-on:click="enterAdding()"><BIconPlus class="icon" /> 新增一筆</button>
            <!--<button type="button" style="margin-right:5px;" class="btn btn-secondary btn-sm" v-on:click="exportJSON()"><BIconBoxArrowInUp class="icon" /> 匯出檔案清單</button>-->
            <label class="btn btn-secondary btn-sm">
              <BIconBoxArrowInDown class="icon" /> 匯入檔案清單(.csv檔)
              <input  type="file" style="display:none;" accept=".csv" @change="onFileChange">
            </label>
          </div>
          <div v-if="currentEditFileRow !== -1">
              <button v-if="currentEditFileRow === -2" style="margin-right:5px;"  type="button" class="btn btn-primary btn-sm" v-on:click="saveAdding()"><BIconCheck class="icon" /> 儲存</button>
              <!--Show when in editing or adding mode-->
              <button v-if="currentEditFileRow === -2 || currentEditFileRow >= 0"  type="button" class="btn btn-secondary btn-sm" v-on:click="backToDefault()"><BIconXLg class="icon" /> 取消</button>
          </div>
        </div>
        <div class="card-body" style="overflow-x:auto;">
        <table class="table table-borderless" style="text-align:center;">
          <thead>
            <tr>
              <th scope="col" class="text-nowrap">#</th>
              <th scope="col" class="text-nowrap">檔案路徑</th>
              <th scope="col" class="text-nowrap">工作表</th>
              <th scope="col" class="text-nowrap">起始列</th>
              <th scope="col"></th>
              <th scope="col"></th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="item,index in files" :key="index">
              <th scope="row">{{index+1}}</th>
              <td>
                <b v-if="currentEditFileRow !== index">{{item.path}}</b>
                <input v-if="currentEditFileRow === index" class="form-control" type="text" v-model="currentPath">
              </td>
              <td>
                <b v-if="currentEditFileRow !== index">{{item.sheet}}</b>
                <input v-if="currentEditFileRow === index" class="form-control" type="text" v-model="currentSheet">
              </td>
              <td>
                <b v-if="currentEditFileRow !== index">{{item.firstrow}}</b>
                <input v-if="currentEditFileRow === index" class="form-control" type="number" min="1" v-model="currentFirstRow">
              </td>
              <td>
                <button v-if="currentEditFileRow === -1" type="button" class="btn btn-secondary btn-sm" v-on:click="editFileRow(index)"><BIconPen class="icon" /></button>
                <button v-if="currentEditFileRow === index" type="button" class="btn btn-primary btn-sm" v-on:click="saveFileRow(index)"><BIconCheck class="icon" /></button>
              </td>
              <td>
                <button v-if="currentEditFileRow === -1" type="button" class="btn btn-danger btn-sm" v-on:click="removeFileRow(index)"><BIconXLg class="icon" /></button>
              </td>
            </tr>
            <tr v-if="currentEditFileRow === -2">
              <th scope="row"></th>
              <td><input  class="form-control" type="text" v-model="currentPath"  placeholder="請輸入Excel檔案路徑"></td>
              <td><input  class="form-control" type="text" v-model="currentSheet" placeholder="請輸入Sheet名稱"></td>
              <td><input  class="form-control" type="number" min="1" v-model="currentFirstRow" ></td>
            </tr>
          </tbody>
        </table>
        </div>
        </div>
      </div>

      <div class="col-4" v-if="currentEditFileRow === -2 || currentEditFileRow >= 0"> <!--Show when in editing or adding mode-->
      <div class="card">
      <div class="card-header d-flex justify-content-between align-items-center">
        <b>編輯欄位</b>
        <button type="button" class="btn btn-secondary btn-sm" v-on:click="copyLastColumn()"><BIconBoxArrowInDown class="icon" /> 複製前一筆欄位</button>
      </div>
      <div class="card-body">
      <table class="table table-borderless" style="text-align:center;">
        <thead>
          <tr>
            <th scope="col">#</th>
            <th scope="col">欄位名稱</th>
            <th scope="col">欄位位置</th>
            <th scope="col"></th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="item,index in currentColumns" :key="index">
            <th scope="row">{{index+1}}</th>
            <td><input class="form-control" type="text" v-model="item.name"></td>
            <td><input class="form-control" type="number" min="1" v-model="item.pos"></td>
            <td><button type="button" class="btn btn-danger btn-sm" v-on:click="removeColumnRow(index)"><BIconXLg class="icon" /></button></td>
          </tr>
          <tr> <!--Show when in editing or adding mode-->
            <th scope="row"></th>
            <td><input class="form-control" type="text" v-model="currentColumnName" placeholder="請輸入欄位名稱"></td>
            <td><input class="form-control" type="number" min="1" v-model="currentColumnPos"></td>
            <td><button type="button" class="btn btn-success btn-sm" v-on:click="addColumnRow()"><BIconPlus class="icon" /></button></td>
          </tr>
        </tbody>
      </table>
      </div>
      </div>
      </div>
    </div>

    <div class="row">
      <div class="col">
        <div class="card" v-if="layout.length > 0" style="margin-top:10px;">
            <div class="card-header d-flex justify-content-between align-items-center"><b>欄位預覽</b></div>
            <div class="card-body" style="overflow-x:auto;">
            <table class="table table-borderless" style="text-align:center;">
                <thead>
                <tr>
                    <th scope="col" class="text-nowrap">檔案</th>
                    <th scope="col" class="text-nowrap" v-for="key in Object.keys(layout[0])" v-bind:key="key">{{key}}</th>
                </tr>
                </thead>
                <tbody>
                <tr v-for="item,index in layout" v-bind:key="index">
                    <th scope="col">{{index+1}}</th>
                    <td v-for="value,index in Object.values(item)" v-bind:key="index">{{value}}</td>
                </tr>
                </tbody>
            </table>
            </div>
        </div>
      </div>
    </div>

  </div>

  <footer class="footer mt-auto">
    <p class="text-muted" style="margin: 0px;text-align: center;">版本：1.0.0</p>
  </footer>

</template>
<script>

//replaceAll Polyfill

/**
 * String.prototype.replaceAll() polyfill
 * https://gomakethings.com/how-to-replace-a-section-of-a-string-with-another-one-with-vanilla-js/
 * @author Chris Ferdinandi
 * @license MIT
 */
if (!String.prototype.replaceAll) {
	String.prototype.replaceAll = function(str, newStr){

		// If a regex pattern
		if (Object.prototype.toString.call(str).toLowerCase() === '[object regexp]') {
			return this.replace(str, newStr);
		}

		// If a string
		return this.replace(new RegExp(str, 'g'), newStr);

	};
}

//Clean Punctuation
String.prototype.clsPunc = function(){
  return this.replace(/[\p{P}\p{S}\p{Z}]/gu, '').toLowerCase()
}


export default {
  name: 'files',
  data(){
    return {
      currentPath:"",
      currentSheet:"",
      currentFirstRow:1,
      currentColumns:[],
      currentColumnName:"",
      currentColumnPos:1,
      // -1 -> not editing or adding
      // >0 -> indicate editing row
      // -2 -> adding
      currentEditFileRow:-1, 
      files:[],
    }
  },
  components: {
    
  },
  watch: {
    files: {
      handler(val) {
        if (typeof window.Alteryx !== 'undefined'){ 
          let files = JSON.stringify(val)
          window.Alteryx.Gui.Manager.getDataItem("files").setValue(files)
        }
      },
      deep: true
    },
  },
  mounted() {
    if (typeof window.Alteryx !== 'undefined'){
      //Load Alteryx Library
      let libpath = window.Alteryx.LibDir+"2/lib/build/designerDesktop.bundle.js"
      let script  = document.createElement('script')
      script.setAttribute('src', libpath)
      //Script Onload Callback
      script.onload = function(){
        //Define DataItem
        window.Alteryx.Gui.BeforeLoad = function (manager, AlteryxDataItems) {
          var files = new AlteryxDataItems.SimpleString('files')
          manager.addDataItem(files)
          var workdir = new AlteryxDataItems.SimpleString('workdir')
          manager.addDataItem(workdir)
        }
        //Load Settings
        window.Alteryx.Gui.AfterLoad = function (manager) {
          //Set WorkflowDirectory
          manager.getDataItem("workdir").setValue("%Engine.WorkflowDirectory%")
          //If there is setting in xml, load into vue.
          if ( manager.getDataItem("files").getValue() ){
            let files =  JSON.parse(manager.getDataItem("files").getValue())
            try {
              files = files.map((file)=>{
                let validated = this.validateRow(file.path,file.sheet,file.firstrow,file.columns)
                return validated
              })
              this.files = files
            } catch (err) {
              alert("XML錯誤："+err)
            }
          } 
        }.bind(this)
      }.bind(this)
      //Load Script
      document.head.appendChild(script)
    }
  },
  computed: {
    unionColumns(){
      return [...new Set([].concat.apply([], this.files.map(x=>x.columns)).map(x=>x.name))]
    },
    layout(){
      return this.files.map((file)=>{
        let temp  = {}
        this.unionColumns.forEach((column) =>{
            let search = file.columns.filter(x=>x.name===column)
            if (search.length === 0){
              temp[column] = "空"
            } else{
              temp[column] = search[0].pos
            }
        })
        return temp
      })
    }
  },
  methods:{
    backToDefault(){
      //Current Field back to Default
      this.currentPath=""
      this.currentSheet=""
      this.currentFirstRow=1
      this.currentColumns=[]
      this.currentColumnName=""
      this.currentColumnPos=1
      //Back to not editing or adding
      this.currentEditFileRow = -1
    },
    editFileRow(index){
      //Copy Value to Current
      this.currentPath = this.files[index].path
      this.currentSheet = this.files[index].sheet
      this.currentFirstRow = this.files[index].firstrow
      this.currentColumns  = JSON.parse(JSON.stringify(this.files[index].columns))
      this.currentEditFileRow = index
    },
    //用來驗證 columns 資料是否正確
    //包含
    //1.未輸入欄位名稱或是欄位位置非正整數
    //2.欄位名稱不得重複
    //3.欄位位置不得重複
    //作用於 saveFileRow(儲存目前編輯欄位)、saveAdding(儲存新增欄位)
    validateColumnData(columns){
      //檢查：未輸入欄位名稱或是欄位位置非正整數
      columns = columns.map((item)=>{
        let name = item.name.clsPunc()
        let pos  = item.pos
        if (name!=="" & pos >= 1 & Number.isInteger(pos)){
          //pass
        } else{
          throw "未輸入欄位名稱或是欄位位置非正整數" 
        }
        return {name:name,pos:pos}
      })
      //檢查Columns名稱獨特性
      let colname = columns.map(x=>x.name)
      let uniq    = [...new Set(colname)]
      if(colname.length === uniq.length){
        //pass
      }else{
        throw "欄位名稱不得重複"
      }
      //檢查Columns位置獨特性
      let colpos = columns.map(x=>x.pos)
      let uniq_pos    = [...new Set(colpos)]
      if(colpos.length === uniq_pos.length){
        //pass
      }else{
        throw "欄位位置不得重複"
      }
      //返回
      return columns
    },
    validateRow(path,sheet,firstrow,columns){
      //Check Path
      if (path !==""){
        //pass
      }else{
          throw "檔案路徑不得為空"
      }
      //Check File is XLSX
      if (path.split('.').pop().toLowerCase() === "xlsx"){
        //pass
      }else{
          throw "檔案路徑僅能接受 XLSX 格式之檔案"
      }
      //Check Sheet
      if (sheet !==""){
        //pass
      }else{
          throw "工作表不得為空"
      }
      //Check Firstrow is positive integer
      if (Number.isInteger(firstrow) & firstrow >=1){
        //pass
      }else{
          throw "起始列應為正整數"
      }
      //Run Columns Check
      columns = this.validateColumnData(columns)
      //There should be at least one column
      if (columns.length>0){
        //pass
      }else{
          throw "無欄位資料，請新增欄位資訊"
      }
      return {path:path,sheet:sheet,firstrow:firstrow,columns:columns}
    },
    saveFileRow(index){
      try{
        let validated = this.validateRow(this.currentPath,this.currentSheet,this.currentFirstRow,JSON.parse(JSON.stringify(this.currentColumns)))
        //Save Back
        this.files[index].path     = validated.path
        this.files[index].sheet    = validated.sheet
        this.files[index].firstrow = validated.firstrow
        this.files[index].columns  = validated.columns
        //Current Field back to Default
        this.backToDefault()
      } catch(err){
        alert(err)
      }
    },
    removeFileRow(index){
      this.files.splice(index,1)
    },
    addColumnRow(){
      let name = this.currentColumnName.clsPunc()
      let pos  = this.currentColumnPos
      this.currentColumnName = ""
      this.currentColumnPos = 1
      try{
        //基本檢查
        if (name !== "" & pos >=1 & Number.isInteger(pos)){
          //pass
        } else {
          throw "未輸入欄位名稱或是欄位位置非正整數"
        }
        //確認目前currentColumns中無相同name
        let num_of_existed = this.currentColumns.filter(x=> x.name===name).length
        if (num_of_existed === 0){
          //pass
        } else {
          throw "欄位名稱不得重複"
        }
        //加入
        this.currentColumns.push({name:name,pos:pos})
      } catch(err){
        alert(err)
      }
    },
    removeColumnRow(index){
      this.currentColumns.splice(index,1)
    },
    enterAdding(){
      this.currentEditFileRow=-2
    },
    saveAdding(){
      try{
        let validated = this.validateRow(this.currentPath,this.currentSheet,this.currentFirstRow,JSON.parse(JSON.stringify(this.currentColumns)))
        //New
        this.files.push({
          path:validated.path,
          sheet:validated.sheet,
          firstrow:validated.firstrow,
          columns:validated.columns
        })
        //Current Field back to Default
        this.backToDefault()
      } catch(err){
        alert(err)
      }
    },
    copyLastColumn(){
      //於不編輯亦非新增模式時
      if(this.currentEditFileRow === -1){
        //pass
      }
      //於新增模式時
      else if(this.currentEditFileRow === -2){
        if (this.files.length===0){
          alert("無前一筆資料")
        } else {
          this.currentColumns = JSON.parse(JSON.stringify(this.files[this.files.length - 1].columns))
        }
      }
      //於編輯模式時
      else{
        if(this.currentEditFileRow === 0){
          alert("無前一筆資料")
        }
        else{
          this.currentColumns = JSON.parse(JSON.stringify(this.files[this.currentEditFileRow - 1].columns))
        }
      }
    },
    onFileChange(e) {
     let files = e.target.files || e.dataTransfer.files;
     if (!files.length) return;
     this.readFile(files[0]);
     e.target.value = null
   },
   readFile(file) {
     let reader = new FileReader();
     reader.onload = function(e){
      
      try{
        //解析CSV
        let mapping = {
            "檔案路徑":"path",
            "工作表":"sheet",
            "起始列":"firstrow",
            "選取欄位":"columnpos",
            "欄位名稱":"columnname"
        }
        let data = this.$papa.parse(e.target.result,{header:true,skipEmptyLines:true,transformHeader:(header)=>{
          if (Object.keys(mapping).indexOf(header) !== -1){
            return mapping[header]
          } else {
            throw `檔案中 ${header} 不是可接受的欄位`
          }
        }})
        //確認每筆資料均有mapping中欄位
        data.data.forEach((element,index)=>{
          let keys = Object.keys(element)
          let cols = Object.values(mapping)
          let result = cols.every(function (element) {
            return keys.includes(element);
          });
          if (result){
            //pass
          }else{
            throw `第 ${index+1} 筆資料未包含必要欄位`
          }
        })
        //轉換資料
        let parsed =  data.data.map((row,index)=>{
          let path       = row.path
          let sheet      = row.sheet
          let firstrow   = parseFloat(row.firstrow) 
          let columnname = row.columnname.split(",")
          let columnpos  = row.columnpos.split(",")
          //確認欄位數量相同
          if (columnname.length === columnpos.length){
            //pass
          }else{
            throw `第 ${index+1} 筆資料欄位位置與欄位名稱數量不符`
          }
          let columns = columnpos.map((pos,index)=>{
            return {"name":columnname[index],"pos":parseFloat(pos)}
          })
          let validated = this.validateRow(path,sheet,firstrow,columns)
          return validated
        })
        this.files = parsed
      } catch (err) {
        alert("檔案錯誤："+err)
      }
     }.bind(this)
     reader.readAsText(file,"big5");
   }
  }
}
</script>

<style>
#app {
  font-family: "Helvetica Neue", Helvetica, Arial, "Microsoft JhengHei", "PingFang TC", "Heiti TC", sans-serif;
  display: flex;
  flex-direction: column;
  height: 100%; 
}
html,body{
  height: 100%;
}
</style>
