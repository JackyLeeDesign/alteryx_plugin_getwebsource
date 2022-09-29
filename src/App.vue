<template>

  <div class="container-fluid">
    <!-- PwC logo -->
    <div class="row">
      <div class="col">
        <div style="margin-top:20px;">
          <img src="./PwC.png" style="width: 100px;">
        </div>
      </div>
    </div>

    <!-- 主要內容 -->
    <div class="row">
      <div class="col">
        <!-- 第一步 -->
        <div class="card" style="margin-top:10px;">
          <div class="card-header d-flex justify-content-between align-items-center"><b>請輸入API連結：</b></div>
          <div class="card-body" style="overflow-x:auto;">
            <label for="exampleFormControlInput1" class="form-label"><b>
                <BIconFiles style="vertical-align:text-top;" class="icon" /> 本元件根據指定API連結，將資料擷取下來。
              </b></label>
            <input type="text" class="form-control" placeholder="API連結" v-model="url">
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
  String.prototype.replaceAll = function (str, newStr) {

    // If a regex pattern
    if (Object.prototype.toString.call(str).toLowerCase() === '[object regexp]') {
      return this.replace(str, newStr);
    }

    // If a string
    return this.replace(new RegExp(str, 'g'), newStr);

  };
}

//Clean Punctuation
String.prototype.clsPunc = function () {
  return this.replace(/[\p{P}\p{S}\p{Z}]/gu, '').toLowerCase()
}


export default {
  name: 'files',
  data() {
    return {
      url: "",
      webData: ""
    }
  },
  components: {

  },
  watch: {
    url: {
      handler(val) {
        if (typeof window.Alteryx !== 'undefined') {
          window.Alteryx.Gui.Manager.getDataItem("url").setValue(val)
        }
      },
      deep: true
    },
  },
  mounted() {
    if (typeof window.Alteryx !== 'undefined') {
      //Load Alteryx Library
      let libpath = window.Alteryx.LibDir + "2/lib/build/designerDesktop.bundle.js"
      let script = document.createElement('script')
      script.setAttribute('src', libpath)
      //Script Onload Callback
      script.onload = function () {
        //Define DataItem
        window.Alteryx.Gui.BeforeLoad = function (manager, AlteryxDataItems) {
          var url = new AlteryxDataItems.SimpleString('url')
          manager.addDataItem(url)
          var workdir = new AlteryxDataItems.SimpleString('workdir')
          manager.addDataItem(workdir)
        }
        //Load Settings
        window.Alteryx.Gui.AfterLoad = function (manager) {
          //Set WorkflowDirectory
          manager.getDataItem("workdir").setValue("%Engine.WorkflowDirectory%")
          this.url = manager.getDataItem("url").getValue()
        }.bind(this)
      }.bind(this)
      //Load Script
      document.head.appendChild(script)
    }
  },
  computed: {
    // unionColumns(){
    //   return [...new Set([].concat.apply([], this.files.map(x=>x.columns)).map(x=>x.name))]
    // },
    // layout(){
    //   return this.files.map((file)=>{
    //     let temp  = {}
    //     this.unionColumns.forEach((column) =>{
    //         let search = file.columns.filter(x=>x.name===column)
    //         if (search.length === 0){
    //           temp[column] = "空"
    //         } else{
    //           temp[column] = search[0].pos
    //         }
    //     })
    //     return temp
    //   })
    // }
  },
  methods: {
    // fetchUrl(url) {

    // }
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

html,
body {
  height: 100%;
}
</style>
