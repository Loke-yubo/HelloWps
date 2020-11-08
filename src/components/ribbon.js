import Util from './js/util.js'
import Formdata from './js/formdata.js'
// import SystemDemo from './js/systemdemo.js'

//记录是否用户点击OA文件的保存按钮
var EnumDocSaveFlag = {
    OADocSave: 1,
    NoneOADocSave: 0
}

//标识文档的落地模式 本地文档落地 0 ，不落地 1
var EnumDocLandMode = {
    DLM_LocalDoc: 0,
    DLM_OnlineDoc: 1
}

//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI){
    if (typeof (wps.ribbonUI) != "object"){
		wps.ribbonUI = ribbonUI
    }
    
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = Util.WPS_Enum
    }

    // //这几个导出函数是给外部业务系统调用的
    // window.openOfficeFileFromSystemDemo = SystemDemo.openOfficeFileFromSystemDemo
    // window.InvokeFromSystemDemo = SystemDemo.InvokeFromSystemDemo
    window.dispatcher = dispatcher;
    window.OnOpenOnLineDocSuccess = OnOpenOnLineDocSuccess;
    window.OnOpenOnLineDocDownFail = OnOpenOnLineDocDownFail;
    window.OnUploadToServerSuccess = OnUploadToServerSuccess;
    window.OnUploadToServerFail =OnUploadToServerFail;

    wps.PluginStorage.setItem("EnableFlag", false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    return true
}

function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            {
                const doc = wps.WpsApplication().ActiveDocument
                if (!doc) {
                    alert("当前没有打开任何文档")
                    return
                }
                alert(doc.Name)
            }
            break;
        case "btnIsEnbable":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                wps.PluginStorage.setItem("EnableFlag", !bFlag)
                
                //通知wps刷新以下几个按饰的状态
                wps.ribbonUI.InvalidateControl("btnIsEnbable")
                wps.ribbonUI.InvalidateControl("btnShowDialog") 
                wps.ribbonUI.InvalidateControl("btnShowTaskPane") 
                //wps.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
                break
            }
        case "btnShowDialog":
            wps.ShowDialog(Util.GetUrlPath() + "dialog", "这是一个对话框网页", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
            break
        case "btnShowTaskPane":
            {
                let tsId = wps.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = wps.CreateTaskPane(Util.GetUrlPath() + "taskpane")
                    let id = tskpane.ID
                    wps.PluginStorage.setItem("taskpane_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = wps.GetTaskPane(tsId)
                    tskpane.Visible = !tskpane.Visible
                }
            }
            break
        case "btnWebNotify":
            {
                let currentTime = new Date()
                let timeStr = currentTime.getHours() + ':' + currentTime.getMinutes() + ":" + currentTime.getSeconds()
                wps.OAAssist.WebNotify("这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:" + timeStr)
            }
            break
        case "btnContent":
            {
                var l_doc = wps.WpsApplication().ActiveDocument;
                alert("保存文档方法：wps.WpsApplication().ActiveDocument.Save()");
                alert("获取文件名：" + l_doc.Name);
                alert("获取文件路径：" + l_doc.FullName);
                alert("获取服务器路径：" + window.document.location.href);

            }
            break
        case "btnSaveToServer":
            {
                OnBtnSaveToServer();
            }
            break
        case "btnSaveAsFile"://另存为本地文件
            {
                OnBtnSaveAsLocalFile();
            }
            break
        default:
            break
    }
    return true
}

/**
 *  执行另存为本地文件操作
 */
function OnBtnSaveAsLocalFile() {

    //检测是否有文档正在处理
    var l_doc = wps.WpsApplication().ActiveDocument;
    if (!l_doc) {
        alert("WPS当前没有可操作文档！");
        return;
    }

    // 设置WPS文档对话框 2 FileDialogType:=msoFileDialogSaveAs
    var l_ksoFileDialog = wps.WpsApplication().FileDialog(2);
    l_ksoFileDialog.InitialFileName = l_doc.Name; //文档名称

    if (l_ksoFileDialog.Show() == -1) { // -1 代表确认按钮
        l_ksoFileDialog.Execute(); //会触发保存文档的监听函数
    }
}

/**
 * web页面调用WPS的方法入口
 *  * info参数结构
 * info:[
 *      {
 *       '方法名':'方法参数',需要执行的方法
 *     },
 *     ...
 *   ]
 * @param {*} info
 */
function dispatcher(info) {

    //执行web页面传递的方法
    for (var index = 0; index < info.length; index++) {
        var func = info[index];
        for (var key in func) {
            if (key === "OpenDoc") {
                OpenDoc(func[key]); //进入打开文档处理函数
            } else if (key === "OnlineEditDoc") { //在线方式打开文档，属于文档不落地的方式打开
                OpenDoc(func[key]);
            } else if (key === "OnlineReadOnlyDoc") { //在线方式打开文档，文档只读，不允许上传到服务器
                OpenOnLineReadOnlyFile(func[key]);
            } else if (key === "getDocumentName") { //获取当前打开的文件名
                getDocumentName();
            }
        }
    }
    return {message:"ok", app:wps.WpsApplication().Name}
}

//获取当前打开的文件名
function getDocumentName(){
    if (wps.WpsApplication().ActiveDocument){
        alert(wps.WpsApplication().ActiveDocument.Name);
    }
}

//打开来自服务器端传递来的文档
function OpenDoc(OaParams) {
    if (OaParams.fileName == "") {
        wps.PluginStorage.setItem("Save2OAShowConfirm", true);
        NewFile(OaParams);
    }else{
        if("" == OaParams.groundDocLandMode || null == OaParams.groundDocLandMode){
            OpenOnLineFile(OaParams, false);
        }else{
            OpenOnLineFile(OaParams, true);
        }
    }
}

/**
 * 作用：文档打开后执行的动作集合
 * @param {*} doc 文档对象
 * @param {*} params 前端传递的参数集合
 * @param {*} isOnlineDoc 在线打开/落地打开
 */
function pOpenFile(doc, params, isOnlineDoc) {
    var l_IsOnlineDoc = isOnlineDoc
    //Office文件打开后，设置该文件属性：从服务端来的文件
    pSetOADocumentFlag(doc, params)
    //设置当前文档为 本地磁盘落地模式
    if (l_IsOnlineDoc == true) {
        DoSetOADocLandMode(doc, EnumDocLandMode.DLM_OnlineDoc);
    } else {
        DoSetOADocLandMode(doc, EnumDocLandMode.DLM_LocalDoc);
    }

    //重新设置工具条按钮的显示状态
    pDoResetRibbonGroups();
    // 触发切换窗口事件
    OnWindowActivate();
    // 把WPS对象置前
    wps.WpsApplication().Activate();
    return doc;
}

/**
 * 打开服务端的文档
 * @param {*} fileUrl 文件url路径
 */
function OpenOnLineFile(params, groundDocLandMode) {
    //参数如果为空的话退出
    if (!params) return;

    //获取在线文档URL
    var l_fileUrl = params.fileName;
    var l_doc;
    if (l_fileUrl && groundDocLandMode) {
        //下载文档不落地（16版WPS的925后支持）
        wps.WpsApplication().Documents.OpenFromUrl(l_fileUrl, "OnOpenOnLineDocSuccess", "OnOpenOnLineDocDownFail");
        l_doc = wps.WpsApplication().ActiveDocument;

        //执行文档打开后的方法
        pOpenFile(l_doc, params, groundDocLandMode);
    }else if(l_fileUrl && !groundDocLandMode){

        //如果当前没有打开文档，则另存为本地文件，再打开
        if (l_fileUrl.indexOf("http") == 0) { // 网络文档
            DownloadFile(l_fileUrl, function (path) {
                if (path == "") {
                    alert("从服务端下载路径：" + l_fileUrl + "\n" + "获取文件下载失败！");
                    return null;
                }
                wps.WpsApplication().Documents.Open(path, false, false, false);
                
                l_doc = wps.WpsApplication().ActiveDocument;

                //执行文档打开后的方法
                pOpenFile(l_doc, params, groundDocLandMode);
            });
        }
    }
    
    return l_doc;
}

/**
 * WPS下载文件到本地打开（业务系统可根据实际情况进行修改）
 * @param {*} url 文件流的下载路径
 * @param {*} callback 下载后的回调
 */
function DownloadFile(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
            //需要业务系统的服务端在传递文件流时，确保请求中的参数有filename
            var fileName = pGetFileName(xhr, url)
            //落地打开模式下，WPS会将文件下载到本地的临时目录，在关闭后会进行清理
            var path = wps.Env.GetTempPath() + "/" + fileName
            var reader = new FileReader();
            reader.onload = function () {
                wps.FileSystem.writeAsBinaryString(path, reader.result);
                callback(path);
            };
            reader.readAsBinaryString(xhr.response);
        }
    }
    xhr.open('GET', url);
    xhr.responseType = 'blob';
    xhr.send();
}

/**
 * 从requst中获取文件名（确保请求中有filename这个参数）
 * @param {*} request 
 * @param {*} url 
 */

function pGetParamName(data, attr) {
    var start = data.indexOf(attr);
    data = data.substring(start + attr.length);
    return data;
}

function pGetFileName(request, url) {
    var disposition = request.getResponseHeader("Content-Disposition");
    var filename = "";
    if (disposition) {
        var matchs = pGetParamName(disposition, "filename=");
        if (matchs) {
            filename = decodeURIComponent(matchs);
        } else {
            filename = "petro" + Date.getTime();
        }
    } else {
        filename = url.substring(url.lastIndexOf("/") + 1);
    }
    return filename;
}

/**
 * 浏览服务端的文档（不落地）
 * @param {*} fileUrl 文件url路径
 */
function OpenOnLineReadOnlyFile(OAParams) {
    //OA参数如果为空的话退出
    if (!OAParams) return;

    //获取在线文档URL
    var l_OAFileUrl = OAParams.fileName;
    var l_doc;
    if (l_OAFileUrl) {
        //下载文档不落地（16版WPS的925后支持）
        wps.WpsApplication().Documents.OpenFromUrl(l_OAFileUrl, "OnOpenOnLineDocSuccess", "OnOpenOnLineDocDownFail");
        l_doc = wps.WpsApplication().ActiveDocument;
    }
    //执行文档打开后的方法
    pOpenFile(l_doc,OAParams,true);
    return l_doc;
}

/**
 * 打开在线文档成功后触发事件
 * @param {*} resp 
 */
function OnOpenOnLineDocSuccess(resp) {
    console.log(resp);
}

/**
 *  打开在线不落地文档出现失败时，给予错误提示
 */
function OnOpenOnLineDocDownFail() {
    alert("打开在线不落地文档失败！请尝试重新打开。");
    return;
}



/**
 * 作用：设置Ribbon工具条的按钮显示状态
 * @param {*} paramsGroups 
 */
function pDoResetRibbonGroups(paramsGroups) {
    console.log(paramsGroups)
}

/**
 * 从OA调用传来的指令，打开本地新建文件
 * @param {*} fileUrl 文件url路径
 */
function NewFile(params) {
    //获取WPS Application 对象
    var wpsApp = wps.WpsApplication();
    wps.PluginStorage.setItem("IsInCurrOADocOpen", true); //设置OA打开文档的临时状态
    var doc = wpsApp.Documents.Add(); //新增OA端文档
    wps.PluginStorage.setItem("IsInCurrOADocOpen", false);

    
    //Office文件打开后，设置该文件属性：从服务端来的OA文件
    pSetOADocumentFlag(doc, params);
    //设置当前文档为 本地磁盘落地模式   
    DoSetOADocLandMode(doc, EnumDocLandMode.DLM_LocalDoc);
    //强制执行一次Activate事件
    OnWindowActivate();
    wps.WpsApplication().Activate(); //把WPS对象置前

    return doc; //返回新创建的Document对象
}

//切换窗口时触发的事件
function OnWindowActivate() {
    var l_doc = wps.WpsApplication().ActiveDocument;
    SetCurrDocEnvProp(l_doc); // 设置当前文档对应的用户名
    showOATab(); // 根据文件是否为系统文件来显示系统菜单再进行刷新按钮
    setTimeout(activeTab, 500); // 激活页面必须要页签显示出来，所以做1秒延迟
    return;
}

function activeTab() {
    //启动WPS程序后，默认显示的工具栏选项卡为ribbon.xml中某一tab
    if (wps.ribbonUI)
        wps.ribbonUI.ActivateTab('WPSWorkExtTab');
}

function showOATab() {
    wps.PluginStorage.setItem("ShowOATabDocActive", pCheckIfOADoc()); //根据文件是否为OA文件来显示OA菜单
    wps.ribbonUI.Invalidate(); // 刷新Ribbon自定义按钮的状态
}

/**
 *  作用：根据当前活动文档的情况判断，当前文档适用的系统参数，例如：当前文档对应的用户名称等
 */
function SetCurrDocEnvProp(doc) {
    if (!doc) return;
    var l_bIsOADoc = false;
    l_bIsOADoc = pCheckIfOADoc(doc);

    //如果是OA文件，则按OA传来的用户名设置WPS   OA助手WPS用户名设置按钮冲突
    if (l_bIsOADoc == true) {
        var l_userName = GetDocParamsValue(doc, "userName");
        if (l_userName != "") {
            wps.WpsApplication().UserName = l_userName;
            return;
        }
    }
    //如果是非OA文件或者参数的值是空值，则按WPS安装默认用户名设置
    wps.WpsApplication().UserName = wps.PluginStorage.getItem("WPSInitUserName");
}

/**
 * 参数
 * doc : 当前OA文档的Document对象
 * DocLandMode ： 落地模式设置
 */
function DoSetOADocLandMode(doc, DocLandMode) {
    if (!doc) return;
    var l_Param = wps.PluginStorage.getItem(doc.DocID);
    var l_objParam = JSON.parse(l_Param);
    //增加属性，或设置
    l_objParam.groundDocLandMode = DocLandMode; //设置文档的落地标志

    var l_p = JSON.stringify(l_objParam);
    //将文档落地模式标志存入系统变量对象保存

    wps.PluginStorage.setItem(doc.DocID, l_p);

}

//Office文件打开后，设置该文件属性：从服务端来的OA文件
function pSetOADocumentFlag(doc, params) {
    if (!doc) {
        return; //
    }

    var l_Param = params;

    if (doc) {
        var l_p = JSON.stringify(l_Param);
        //将OA文档标志存入系统变量对象保存
        wps.PluginStorage.setItem(doc.DocID, l_p);
    }
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        case "btnContent":
            return "images/5.svg"
        case "btnSaveToServer":
            return "images/w_Save.png"
        case "btnSaveAsFile": //另存为本地文件
            return "images/w_SaveAs.png";
        default:
    }
    return "images/newFromTemp.svg"
}

function OnGetEnabled(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return true
        case "btnShowDialog":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
            }
        case "btnShowTaskPane":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                return bFlag
            }
        default:
            break
    }
    return true
}

function OnGetVisible(control){
    const eleId = control.Id
    console.log(eleId)
    return true
}

function OnGetLabel(control){
    const eleId = control.Id
    switch (eleId) {
    case "btnIsEnbable":
        {
            let bFlag = wps.PluginStorage.getItem("EnableFlag")
            return bFlag ?  "按钮Disable" : "按钮Enable"
        }
    }
    return ""
}

//判断当前文档是否是系统文档
function pCheckIfOADoc() {
    var doc = wps.WpsApplication().ActiveDocument;
    if (!doc){
        return false;
    }else{
        return true;
    }  
}

/**
 *  作用：判断当前文档是否是只读文档
 *  返回值：布尔
 */
function pISOADocReadOnly(doc) {
    if (!doc) {
        return false;
    }
    var l_openType = GetDocParamsValue(doc, "openType"); // 获取传入的参数 openType
    if (l_openType == "") {
        return false;
    }
    try {
        if (l_openType.protectType != -1) { // -1 为未保护
            return true;
        }
    } catch (err) {
        return false;
    }
}

/**
 * 根据传入Document对象，获取OA传入的参数的某个Key值的Value
 * @param {*} Doc 
 * @param {*} Key 
 * 返回值：返回指定 Key的 Value
 */
function GetDocParamsValue(Doc, Key) {
    if (!Doc) {
        return "";
    }

    var l_Params = wps.PluginStorage.getItem(Doc.DocID);
    if (!l_Params) {
        return "";
    }

    var l_objParams = JSON.parse(l_Params);
    if (typeof (l_objParams) == "undefined") {
        return "";
    }

    var l_rtnValue = l_objParams[Key];
    if (typeof (l_rtnValue) == "undefined" || l_rtnValue == null) {
        return "";
    }
    return l_rtnValue;
}

/**
 * 作用：判断是否是不落地文档
 * 参数：doc 文档对象
 * 返回值： 布尔值
 */

function pIsOnlineOADoc(doc) {
    var l_LandMode = GetDocParamsValue(doc, "groundDocLandMode"); //获取文档落地模式
    if (l_LandMode == "") { //用户本地打开的文档
        return false;
    }
    return l_LandMode == EnumDocLandMode.DLM_OnlineDoc;
}

//保存到OA后台服务器
function OnBtnSaveToServer() {
    // console.log('SaveToServer');
    var l_doc = wps.WpsApplication().ActiveDocument;
    if (!l_doc) {
        alert("空文档不能保存！");
        return;
    }

    //非系统打开的文档，不能直接上传到系统
    if (pCheckIfOADoc() == false) {
        alert("非系统打开的文档，不能直接上传到系统！");
        return;
    }

    //如果是OA打开的文档，并且设置了保护的文档，则不能再上传到OA服务器
    if (pISOADocReadOnly(l_doc)) {
        wps.alert("系统设置了保护的文档，不能再提交到系统后台。");
        return;
    }

    /**
     * 参数定义：OAAsist.UploadFile(name, path, url, field,  "OnSuccess", "OnFail")
     * 上传一个文件到远程服务器。
     * name：为上传后的文件名称；
     * path：是文件绝对路径；
     * url：为上传地址；
     * field：为请求中name的值；
     * 最后两个参数为回调函数名称；
     */
    var l_uploadPath = GetDocParamsValue(l_doc, "uploadPath"); // 文件上载路径
    if (l_uploadPath == "") {
        wps.alert("系统未传入文件上载路径，不能执行上传操作！");
        return;
    }

    var l_showConfirm = wps.PluginStorage.getItem("Save2OAShowConfirm")
    if (l_showConfirm) {
        if (!wps.confirm("先保存文档，并开始上传到系统后台，请确认？")) {
            return;
        }
    }

    var l_FieldName = GetDocParamsValue(l_doc, "uploadFieldName"); //上载到后台的业务方自定义的字段名称
    if (l_FieldName == "") {
        l_FieldName = wps.PluginStorage.getItem("DefaultUploadFieldName"); // 默认为‘file’
    }

    var l_UploadName = GetDocParamsValue(l_doc, "uploadFileName"); //设置传入的文件名称参数
    if (l_UploadName == "") {
        l_UploadName = l_doc.Name; //默认文件名称就是当前文件编辑名称
    }

    var l_DocPath = l_doc.FullName; // 文件所在路径

    if (pIsOnlineOADoc(l_doc) == false) {
        //对于本地磁盘文件上传OA，先用Save方法保存后，再上传
        //设置用户保存按钮标志，避免出现禁止OA文件保存的干扰信息
        wps.PluginStorage.setItem("OADocUserSave", EnumDocSaveFlag.OADocSave);
        l_doc.Save(); //执行一次保存方法
        //设置用户保存按钮标志
        wps.PluginStorage.setItem("OADocUserSave", EnumDocSaveFlag.NoneOADocSave);
        //落地文档，调用UploadFile方法上传到OA后台
        try {
            //调用OA助手的上传方法
            UploadFile(l_UploadName, l_DocPath, l_uploadPath, l_FieldName, OnChangeSuffixUploadSuccess, OnChangeSuffixUploadFail);
        } catch (err) {
            alert("上传文件失败！请检查系统上传参数及网络环境！");
        }
    } else {
        if(wps.WpsApplication().BuildFull.indexOf(".8.") > -1){//企业版
            // 不落地的文档，调用 Document 对象的不落地上传方法
            wps.PluginStorage.setItem("OADocUserSave", EnumDocSaveFlag.OADocSave);
            try {
                //调用不落地上传方法
                l_doc.SaveAsUrl(l_UploadName, l_uploadPath, l_FieldName, "OnUploadToServerSuccess", "OnUploadToServerFail");
            } catch (err) {
                alert("上传文件失败！请检查系统上传参数及网络环境，重新上传。");
            }
            wps.PluginStorage.setItem("OADocUserSave", EnumDocSaveFlag.NoneOADocSave);
        }else if(wps.WpsApplication().BuildFull.indexOf(".1.") > -1){//个人版
            alert("上传失败！个人版暂不支持此功能，请下载企业版重试。");
        }
        
    }

    //获取OA传入的 转其他格式上传属性
    var l_suffix = GetDocParamsValue(l_doc, "suffix");
    if (l_suffix == "") {
        console.log("上传需转换的文件后缀名错误，无法进行转换上传!");
        return;
    }

    //判断是否同时上传PDF等格式到OA后台
    var l_uploadWithAppendPath = GetDocParamsValue(l_doc, "uploadWithAppendPath"); //标识是否同时上传suffix格式的文档
    if (l_uploadWithAppendPath == "1") {
        //调用转pdf格式函数，强制关闭转换修订痕迹，不弹出用户确认的对话框
        //pDoChangeToOtherDocFormat(l_doc, l_suffix, false, false);
    }
    return;
}

/**
 * 调用文件上传到服务端时，
 * @param {*} resp 
 */
function OnUploadToServerSuccess(resp) {
    console.log(resp);
    var l_doc = wps.WpsApplication().ActiveDocument;
    // var l_showConfirm = wps.PluginStorage.getItem("Save2OAShowConfirm");
    // if (l_showConfirm) {
    if (wps.confirm("文件上传成功！继续编辑请确认，取消关闭文档。") == false) {
        if (l_doc) {
            console.log("OnUploadToServerSuccess: before Close");
            l_doc.Close(-1); //保存文档后关闭
            console.log("OnUploadToServerSuccess: after Close");
        }
    }
    // }

    // var l_NofityURL = GetDocParamsValue(l_doc, constStrEnum.notifyUrl);
    // if (l_NofityURL != "") {
    //     l_NofityURL = l_NofityURL.replace("{?}", "2"); //约定：参数为2则文档被成功上传
    //     NotifyToServer(l_NofityURL);
    // }
}

function OnUploadToServerFail(resp) {
    alert("文件上传失败！错误信息：" + resp);
}



/**
 * 作用：转格式保存上传成功后，触发这个事件的回调
 * @param {} response 
 */
function OnChangeSuffixUploadSuccess(response) {
    handleResultBody(response);
    var l_doc = wps.WpsApplication().ActiveDocument;
    if (wps.confirm("文件上传成功！继续编辑请确认，取消关闭文档。") == false) {
        if (l_doc) {
            var tempFilePath = l_doc.FullName;
            l_doc.Close(-1); //保存文档后关闭
            afterClose(tempFilePath);
        }
    }
}

function afterClose(path){
    //删除临时文件
    wps.FileSystem.Remove(path);
}

/**
 * 作用：转格式保存失败，触发失败事件回调
 * @param {*} response 
 */
function OnChangeSuffixUploadFail(response) {
    var l_result = "";
    l_result = handleResultBody(response);
    alert("保存失败" + "\n" + +"系统返回数据：" + +JSON.stringify(l_result));
}


/**
 * 解析返回response的参数
 * @param {*} resp 
 * @return {*} body
 */
function handleResultBody(resp) {
    var l_result = "";
    if (resp.Body) {
        //解析返回response的参数
    }
    return l_result;
}



/**
 * WPS上传文件到服务端（业务系统可根据实际情况进行修改，为了兼容中文，服务端约定用UTF-8编码格式）
 * @param {*} strFileName 上传到服务端的文件名称（包含文件后缀）
 * @param {*} strPath 上传文件的文件路径（文件在操作系统的绝对路径）
 * @param {*} uploadPath 上传文件的服务端地址
 * @param {*} strFieldName 业务调用方自定义的一些内容可通过此字段传递，默认赋值'file'
 * @param {*} OnSuccess 上传成功后的回调
 * @param {*} OnFail 上传失败后的回调
 */
function UploadFile(strFileName, strPath, uploadPath, strFieldName, OnSuccess, OnFail) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', uploadPath);

    var fileData = wps.FileSystem.readAsBinaryString(strPath);
    // eslint-disable-next-line no-undef
    var data = new FakeFormData();
    if (strFieldName == "" || typeof strFieldName == "undefined"){//如果业务方没定义，默认设置为'file'
        strFieldName = 'file';
    }
    data.append(strFieldName, {
        name: utf16ToUtf8(strFileName), //主要是考虑中文名的情况，服务端约定用utf-8来解码。
        type: "application/octet-stream",
        getAsBinary: function () {
            return fileData;
        }
    });
    xhr.onreadystatechange = function () {
        if (xhr.readyState == 4) {
            if (xhr.status == 200)
                OnSuccess(xhr.response)
            else
                OnFail(xhr.response);
        }
    };
    xhr.setRequestHeader("Cache-Control", "no-cache");
    xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");
    if (data.fake) {
        xhr.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + data.boundary);
        var arr = StringToUint8Array(data.toString());
        xhr.send(arr);
    } else {
        xhr.send(data);
    }
}

function StringToUint8Array(string) {
    var binLen, buffer, chars;
    binLen = string.length;
    buffer = new ArrayBuffer(binLen);
    chars = new Uint8Array(buffer);
    for (var i = 0; i < binLen; ++i) {
        chars[i] = String.prototype.charCodeAt.call(string, i);
    }
    return buffer;
}

//UTF-16转UTF-8
function utf16ToUtf8(s) {
    if (!s) {
        return;
    }
    var i, code, ret = [],
        len = s.length;
    for (i = 0; i < len; i++) {
        code = s.charCodeAt(i);
        if (code > 0x0 && code <= 0x7f) {
            //单字节
            //UTF-16 0000 - 007F
            //UTF-8  0xxxxxxx
            ret.push(s.charAt(i));
        } else if (code >= 0x80 && code <= 0x7ff) {
            //双字节
            //UTF-16 0080 - 07FF
            //UTF-8  110xxxxx 10xxxxxx
            ret.push(
                //110xxxxx
                String.fromCharCode(0xc0 | ((code >> 6) & 0x1f)),
                //10xxxxxx
                String.fromCharCode(0x80 | (code & 0x3f))
            );
        } else if (code >= 0x800 && code <= 0xffff) {
            //三字节
            //UTF-16 0800 - FFFF
            //UTF-8  1110xxxx 10xxxxxx 10xxxxxx
            ret.push(
                //1110xxxx
                String.fromCharCode(0xe0 | ((code >> 12) & 0xf)),
                //10xxxxxx
                String.fromCharCode(0x80 | ((code >> 6) & 0x3f)),
                //10xxxxxx
                String.fromCharCode(0x80 | (code & 0x3f))
            );
        }
    }

    return ret.join('');

}



//这些函数是给wps客户端调用的
export default {
    OnAddinLoad,
    OnAction,
    GetImage,
    OnGetEnabled,
    OnGetVisible,
    OnGetLabel,
    dispatcher,
    Formdata
};



