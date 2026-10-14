# AI 文件整理助手

用 DeepSeek AI 自动给杂乱文件夹分类、归档的 Windows 桌面小工具。双击 exe 后通过网页界面操作，无需任何技术背景。

## 功能

**第一步：AI 分析 → 生成 Excel 总表**
- 扫描指定文件夹里的 Word / PDF / PPT / TXT 文件，提取文字内容
- 每个文件由 AI 产出 3 个关键词 + 一句话总结 + 建议分类
- 再通盘整个文件库设计统一分类体系（5–12 个互斥分类），保证同类文件归到一起
- 生成《文件整理总表.xlsx》：矩阵布局，每个分类一列、每个文件一个格子，点击格子直接打开原文件

**第二步：按总表归档（可撤销）**
- 在总表里手动调整分类（剪切格子到别的列，或改列表头）
- 归档后文件移动到对应分类的子文件夹，重名自动加序号，表格链接自动更新
- 每次归档写 undo.json 日志，可一键撤销全部还原

**设计亮点**
- 按文件内容 MD5 缓存，重复运行不重复消耗 API 费用；分析失败不落缓存，下次自动重试
- 手动改过的分类重跑时永远保留，不会被 AI 覆盖
- 文件库有增删改时自动重新全局分类
- API Key 只存在本机 config.json，不经过任何第三方服务器

## 快速开始（直接使用）

1. 从 [Releases](../../releases) 下载 `AI文件整理助手_分发版.zip`，解压
2. 双击 `AI文件整理助手.exe`（如弹 SmartScreen 警告：更多信息 → 仍要运行）
3. 在打开的网页界面填写：要整理的文件夹路径 + 你的 DeepSeek API 密钥（[platform.deepseek.com](https://platform.deepseek.com) 注册创建）
4. 点开始分析，完成后生成《文件整理总表.xlsx》

## 从源码运行

```bash
pip install openai openpyxl markitdown
python app.py        # 启动网页界面
```

命令行方式（可选）：

```bash
cp config.example.json config.json   # 填入 api_key 和 source_folder
python main.py                       # 分析并生成总表
python organize.py                   # 按总表归档
python organize.py undo              # 撤销上一次归档
```

## 打包 exe

```bash
pip install pyinstaller
pyinstaller AI文件整理助手.spec
```

## 说明

- 仅支持 Windows 10/11（64 位），需联网调用 DeepSeek API
- 支持 .docx / .pdf / .pptx / .txt；老版 .doc 请先另存为 .docx
- 分析消耗你自己 API Key 的额度（几十上百个文件约几分到几毛钱）

## License

MIT
