# 表格提取 extractTab

把 PDF 里的表格提取成 Excel。Windows 桌面小工具，不用安装，双击即用。

> 本仓库放的是**源码**，不含编译好的 exe。可执行文件请到 [Releases](../../releases) 下载，或按下面的「从源码运行」自行构建。

## 怎么用

1. 双击 `extractTab.exe`
2. 把 PDF 拖进左侧虚线框，或点「添加 PDF」
3. 程序自动找出所有表格，列在中间
4. 勾选要的（默认全选），点「导出 Excel」

一篇论文通常有 3–5 个表，导出后是一个工作簿，**每个表一个工作表**，表名取自原文的表号（如 `Table 2`）。

**同时放进多篇论文时**，工作簿第一页是「目录」，列出每个工作表来自哪个文件、第几页、表几、几行几列、有没有可疑值，点表名可以直接跳过去。工作表名也会自动带上文件名，比如 `pone 0040382 Table 2`，不会和另一篇的 `Table 2` 混在一起。

界面里的表格列表同样按文件分组，看着就知道哪个表属于哪篇。

## 关于黄色的格子

论文 PDF 里有一类隐蔽的编码错误：`±` 号常常被保存成数字 `6`，于是 `0.05 ± 0.01` 提取出来是 `0.0560.01`。

程序会自动还原其中的 `±`，但凡是**被改过、或无法确定的格子都会标黄**，Excel 里也带批注。

还有一类是整段字符映射错位，比如负号变成 `{`、小数点变成 `:`。这种没法可靠地自动纠正，程序会**保留原样并标黄**，不会自作主张改成看着像的数字。

**标黄的格子请对照原文核一眼再用。** 尤其是要拿去算平均值、画图的时候。

## 从源码运行

需要 Python 3.9 以上：

```bash
pip install -r requirements.txt
python main.py
```

打包成单个 exe：

```bash
pip install pyinstaller
pyinstaller extractTab.spec
```

产物是 `dist/extractTab.exe`，约 100MB——PyMuPDF、openpyxl 和 Python 运行时都塞在里面了。

自检脚本会把某个目录下的 PDF 全部跑一遍（提取 → 打印统计 → 导出 Excel），用来确认改动没有破坏解析：

```bash
mkdir samples && cp 你的测试.pdf samples/
python verify.py
```

仓库里**没有**附带测试 PDF，`samples/` 已在 `.gitignore` 中排除。

## 项目结构

```
main.py              启动入口
extracttab/
  core.py            表格识别与单元格清洗（± 号还原、错位字符检测）
  export.py          导出 Excel：目录页、批注、可疑值标黄
  api.py             前端 JS 与 Python 之间的桥接
  gui.py             窗口初始化
  ui.html            界面（HTML/CSS/JS）
extractTab.spec      PyInstaller 打包配置
app.ico              程序图标
verify.py            解析自检脚本
```

## 其他

- 想从资源管理器直接调起：右键 PDF →「发送到」→ 放进 SendTo 的快捷方式
- 「附来源列」勾上后，每行前面会多出文件名、页码、表号三列，方便多篇论文汇总后做透视表
- 解析速度大约是每页半秒，几十页的综述会等十几秒

## 发给同事前要知道的

只要发 `extractTab.exe` 这一个文件就行，不用带别的，也不用安装。但它不是「任何电脑都能跑」：

- **系统要 Windows 10 / 11，64 位。** 32 位系统、Windows 7 跑不了（缺 WebView2）
- **首次打开会慢几秒**，程序要先把自己解压到临时目录，之后就快了
- 第一次运行 Windows 可能弹「已保护你的电脑」，点「更多信息」→「仍要运行」。没有代码签名证书，这个提示绕不过去
- 需要系统自带 WebView2（Win11 都有；Win10 一般也有，没有的话装一个 Microsoft Edge WebView2 Runtime）
- 建议先在一台同事的电脑上试一下，确认公司安全软件不拦

### 它会在电脑上留下什么

- 使用过程中会解压到 `%TEMP%\_MEIxxxx`，约 180MB，**正常关窗口会自动删掉**
- 在 `%LOCALAPPDATA%\extractTab` 留约 30MB 浏览器缓存
- 不写注册表，不装服务。想彻底清干净，把上面那个 `extractTab` 文件夹删掉即可

如果同事是用任务管理器强杀进程关掉的，临时目录会残留。正常点窗口右上角的 × 关闭就行。

