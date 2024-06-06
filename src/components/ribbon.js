import Util from "./js/util.js";
import SystemDemo from "./js/systemdemo.js";
import getI18n from "../i18n";
import Api from "./js/api.js";
import Excel from "./js/excel.js";

//源始语言选项，与ribbon.xml中的sourceLangDropDown的item顺序对应
const SrcLang = ["zh-cn", "en"];
//目标语言选项，与ribbon.xml中的targetLangDropDown的item顺序对应
const TgtLang = ["en", "zh-cn"];
//语言领域选项，与ribbon.xml中的domainDropDown的item顺序对应
const Domain = ["general", "law", "finance"];
//i18n选项，与ribbon.xml中的dropDown_UIlang的item顺序对应
const I18n = ["zh-ui", "en-ui"];
//xml中node的id
const Id = [
  "btnReplace",
  "btnAfter",
  "btnReplaceRight",
  "btnReplaceUnder",
  "sourceLangDropDown",
  "targetLangDropDown",
  "domainDropDown",
  "dropDown_UIlang",
  "zh-ui",
  "en-ui",
  "zh-cn",
  "en",
  "general",
  "law",
  "finance",
  "tab1",
];

//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
  if (typeof wps.ribbonUI != "object") {
    wps.ribbonUI = ribbonUI;
  }

  if (typeof wps.Enum != "object") {
    // 如果没有内置枚举值
    wps.Enum = Util.WPS_Enum;
  }

  //这几个导出函数是给外部业务系统调用的
  window.openOfficeFileFromSystemDemo = SystemDemo.openOfficeFileFromSystemDemo;
  window.InvokeFromSystemDemo = SystemDemo.InvokeFromSystemDemo;

  wps.PluginStorage.setItem("SrcLang", SrcLang[0]); //原语言
  wps.PluginStorage.setItem("TgtLang", TgtLang[0]); //目标语言
  wps.PluginStorage.setItem("Domain", Domain[0]); //领域
  wps.PluginStorage.setItem("I18n", I18n[0]); // ui界面语言
  return true;
}

async function translate(callback) {
  if (typeof callback !== "function") {
    throw new Error("callback must be a function");
  }
  const srcLang = wps.PluginStorage.getItem("SrcLang");
  const tgtLang = wps.PluginStorage.getItem("TgtLang");
  const domain = wps.PluginStorage.getItem("Domain");
  try {
    const word = Excel.GetSeletionWord();
    // return console.log(word);
    for (let index = 0; index < word.length; index++) {
      const item = word[index];
      let promise;
      if (item.isTranslate) {
        promise = Api.getLang(item.text).then((res) => {
          if (res.lang === srcLang) {
            return Api.translate(item.text, srcLang, tgtLang, domain);
          } else {
            return Promise.reject();
          }
        });
      } else {
        promise = Promise.resolve(item.text);
      }
      await promise
        .then((result) => {
          item.translate = result;
          // return callback(item, index === 0 ? null : word[index - 1], index === word.length - 1 ? null : word[index + 1])
          return callback(item);
        })
        .catch(() => {
          item.isTranslate = false;
          // return callback(item, index === 0 ? null : word[index - 1], index === word.length - 1 ? null : word[index + 1])
          return callback(item);
        });
    }
  } catch (e) {
    window.Application.alert(e.message);
  }
}

let loading = false;

function OnAction(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "btnReplace":
      {
        if (loading) {
          alert("正在翻译中，请稍等...");
          break;
        }
        loading = true;
        wps.ribbonUI.InvalidateControl("btnReplace");
        wps.ribbonUI.InvalidateControl("btnAfter");
        wps.ribbonUI.InvalidateControl("btnReplaceRight");
        wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        translate(Excel.ReplaceWord).finally(() => {
          loading = false;
          wps.ribbonUI.InvalidateControl("btnReplace");
          wps.ribbonUI.InvalidateControl("btnAfter");
          wps.ribbonUI.InvalidateControl("btnReplaceRight");
          wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        });
      }
      break;
    case "btnAfter":
      {
        if (loading) {
          alert("正在翻译中，请稍等...");
          break;
        }
        loading = true;
        wps.ribbonUI.InvalidateControl("btnReplace");
        wps.ribbonUI.InvalidateControl("btnAfter");
        wps.ribbonUI.InvalidateControl("btnReplaceRight");
        wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        translate(Excel.InsertAfter).then(() => {
          loading = false;
          wps.ribbonUI.InvalidateControl("btnReplace");
          wps.ribbonUI.InvalidateControl("btnAfter");
          wps.ribbonUI.InvalidateControl("btnReplaceRight");
          wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        });
      }
      break;
    case "btnReplaceRight":
      {
        if (loading) {
          alert("正在翻译中，请稍等...");
          break;
        }
        const excel = wps.EtApplication().ActiveWorkbook;
        const selection = wps.EtApplication().Selection;
        let userConfirmed = true;
        if (!selection) {
          alert("没有选中任何单元格");
        }
        const selectedCells = selection.Cells;
        if (selectedCells.Columns.Count > 1) {
          alert("该操作适用于选择区域为一列时。");
          break;
        }
        for (
          let row = selection.Row;
          row <= selection.Row + selectedCells.Rows.Count - 1;
          row++
        ) {
          const cell = excel.ActiveSheet.Cells.Item(row, selection.Column + 1);
          if (cell.Value2 && cell.Value2.trim() !== "") {
            userConfirmed = wps.confirm(
              "检测到选中区右侧单元格有内容，是否继续？"
            );
            break;
          }
        }
        if (!userConfirmed) break;
        loading = true;
        wps.ribbonUI.InvalidateControl("btnReplace");
        wps.ribbonUI.InvalidateControl("btnAfter");
        wps.ribbonUI.InvalidateControl("btnReplaceRight");
        wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        translate(Excel.InsertRight).then(() => {
          loading = false;
          wps.ribbonUI.InvalidateControl("btnReplace");
          wps.ribbonUI.InvalidateControl("btnAfter");
          wps.ribbonUI.InvalidateControl("btnReplaceRight");
          wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        });
      }
      break;
    case "btnReplaceUnder":
      {
        if (loading) {
          alert("正在翻译中，请稍等...");
          break;
        }
        const excel = wps.EtApplication().ActiveWorkbook;
        const selection = wps.EtApplication().Selection;
        let userConfirmed = true;
        if (!selection) {
          alert("没有选中任何单元格");
        }
        const selectedCells = selection.Cells;
        if (selectedCells.Rows.Count > 1) {
          alert("该操作适用于选择区域为一行时。");
          break;
        }
        for (
          let col = selection.Column;
          col <= selection.Column + selectedCells.Columns.Count - 1;
          col++
        ) {
          const cell = excel.ActiveSheet.Cells.Item(selection.Row + 1, col);
          if (cell.Value2 && cell.Value2.trim() !== "") {
            userConfirmed = wps.confirm(
              "检测到选中区下方单元格有内容，是否继续？"
            );
            break;
          }
        }
        if (!userConfirmed) break;
        loading = true;
        wps.ribbonUI.InvalidateControl("btnReplace");
        wps.ribbonUI.InvalidateControl("btnAfter");
        wps.ribbonUI.InvalidateControl("btnReplaceRight");
        wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        translate(Excel.InsertUnder).then(() => {
          loading = false;
          wps.ribbonUI.InvalidateControl("btnReplace");
          wps.ribbonUI.InvalidateControl("btnAfter");
          wps.ribbonUI.InvalidateControl("btnReplaceRight");
          wps.ribbonUI.InvalidateControl("btnReplaceUnder");
        });
      }
      break;
    default:
      break;
  }
  return true;
}

function OnSoure(index) {
  wps.PluginStorage.setItem("SrcLang", SrcLang[index]);
  wps.PluginStorage.setItem("TgtLang", TgtLang[index]);
  wps.ribbonUI.InvalidateControl("sourceLangDropDown");
  wps.ribbonUI.InvalidateControl("targetLangDropDown");
  return true;
}

function OnTarget(index) {
  wps.PluginStorage.setItem("SrcLang", SrcLang[index]);
  wps.PluginStorage.setItem("TgtLang", TgtLang[index]);
  return true;
}

function OnDomain(index) {
  wps.PluginStorage.setItem("Domain", Domain[index]);
  return true;
}

function OnI18n(index) {
  wps.PluginStorage.setItem("I18n", I18n[index]);
  Id.forEach((id) => wps.ribbonUI.InvalidateControl(id));
  return true;
}

function GetImage(control) {
  const eleId = control.Id;
  if (loading) {
    return "images/loading.svg";
  }
  switch (eleId) {
    case "btnReplaceRight":
      return "images/right-arrow.svg";
    case "btnReplaceUnder":
      return "images/down-arrow.svg";
    default:
  }
  return "images/cloud.svg";
}

function Selected(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "sourceLangDropDown":
      return wps.PluginStorage.getItem("SrcLang");
    case "targetLangDropDown":
      return wps.PluginStorage.getItem("TgtLang");
    case "domainDropDown":
      return wps.PluginStorage.getItem("Domain");
    case "dropDown_UIlang":
      return wps.PluginStorage.getItem("I18n");
  }
}

function OnGetLabel(control) {
  const eleId = control.Id;
  return getI18n(wps.PluginStorage.getItem("I18n"), eleId);
}

//这些函数是给wps客户端调用的
export default {
  OnAddinLoad,
  OnAction,
  GetImage,
  OnGetLabel,
  Selected,
  OnSoure,
  OnTarget,
  OnDomain,
  OnI18n,
};
