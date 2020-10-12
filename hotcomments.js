// Variables used by Scriptable.
// These must be at the very top of the file. Do not edit.
// icon-color: brown; icon-glyph: magic;
// author:Nicolas-kings
const res = await getData();
// 初始化组件ui
let widget = await createWidget(res)
// const url = "https://area.sinaapp.com/bingImg/"   //使用必应壁纸作为背景
// const url = res.music_pic   //使用歌曲封面作为背景
const url = "http://p1.music.126.net/uarVFKgUlrI9Z1nr-50cAw==/109951162843608471.jpg"     //固定一张图片,这里我选用城南花已开的封面,图片不能太大，容易崩溃
const i = await new Request(url);
const img = await i.loadImage();
widget.backgroundImage = img
// 如果不是在组件执行，则显示预览
if (!config.runsInWidget) {
  await widget.presentMedium()
}
// 设置桌面组件
Script.setWidget(widget)
Script.complete()

// 创建组件
async function createWidget() {
  let w = new ListWidget()
  w.backgroundColor = new Color("#222222", 1)
  //   标题
  let title = w.addText("\u7f51\u6291\u4e91\u70ed\u8bc4")
  title.font = Font.boldMonospacedSystemFont(20)
  title.textColor = Color.white()
  title.url = "orpheuswidget://"
  w.addSpacer(8)

  //   内容
  let body = w.addText(`❝${res.comments}❞ `)
  body.font = Font.lightMonospacedSystemFont(18)
  body.textColor = Color.white()
  body.textOpacity = 0.88
  body.url = "orpheus://song/" + GetRequest(res.music_url).split(".")[0]
  w.addSpacer(8)


  let foot = w.addText(`—— 评论来自歌曲「${res.name}」`);
  foot.font = Font.lightMonospacedSystemFont(12)
  foot.textColor = Color.orange();
  foot.textOpacity = 0.88;
  foot.rightAlignText();
  
  return w
}

async function getData() {
  const url = 'https://api.66mz8.com/api/music.163.php'
  const request = new Request(url)
  const res = await request.loadJSON()
  
  return res
}

async function GetRequest(url) {
  if (url.indexOf("?") != -1) {
    var str = url.substr(1);
    strs = str.split("=");
	
    return strs[1];
  }
}
