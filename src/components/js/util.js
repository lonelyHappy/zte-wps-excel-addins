//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft:0,
    msoCTPDockPositionRight:2
}

const BASE_API = {
    // 开发
    dev: "http://218.85.119.93:28088",
    // 预发布
    pre: "https://translate.zte.com.cn:28085",
    // 生产
    prod: "https://translate.zte.com.cn:28085",
}

function GetUrlPath() {
    let e = document.location.toString()
    return -1!=(e=decodeURI(e)).indexOf("/")&&(e=e.substring(0,e.lastIndexOf("/"))),e
}

export default{
    WPS_Enum,
    GetUrlPath,
    BASE_API
}