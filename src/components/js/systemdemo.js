
function openOfficeFileFromSystemDemo(param){
    alert("从业务系统传过来的参数为：" + param)
    let jsonObj = JSON.parse(param)
    return {wps加载项项返回: jsonObj.filepath + ", 这个地址给的不正确"}
}

function InvokeFromSystemDemo(param){
    let jsonObj = JSON.parse(param)
    let handleInfo = jsonObj.Index
    switch (handleInfo){
        case "getDocumentName":{
            let docName = ""
            if (wps.WpsApplication().ActiveDocument){
                docName = wps.WpsApplication().ActiveDocument.Name
            }

            return {当前打开的文件名为:docName}
        }

        case "newDocument":{
            let newDocName=""
            let doc = wps.WpsApplication().Documents.Add()
            newDocName = doc.Name

            return {操作结果:"新建文档成功，文档名为：" + newDocName}
        }
    }

    return {其它xxx:""}
}

export default{
    openOfficeFileFromSystemDemo,
    InvokeFromSystemDemo
}