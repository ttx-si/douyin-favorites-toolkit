# 抖音收藏夹提取与口述逐字稿工具箱 (Douyin Favorites Toolkit)

开箱即用的抖音个人收藏夹自动化归档套件。自动化完成**收藏扫描 -> 原始元数据拉取 -> Excel/JSON 结构化 -> 音频内存流抽取 -> Whisper 逐字稿 -> RapidOCR 图文识别 -> Markdown 知识库排版**。

---

## 3 步极速使用 (Quick Start)

### 第 1 步：安装 Python 依赖
```bash
pip install -r requirements.txt
```

### 第 2 步：启动调试浏览器并登录
- **方式 A（双击脚本）**：双击运行 `start_chrome.bat`
- **方式 B（命令行手动启动）**：
  ```bash
  chrome.exe --remote-debugging-port=9223 --user-data-dir="%USERPROFILE%\.douyin-scraper-chrome-profile" "https://www.douyin.com"
  ```
> **说明**：在弹出的 Chrome 中扫码登录抖音账号，并点击进入自己的 **“喜欢与收藏 -> 我的收藏”** 页面。

### 第 3 步：运行提取与转录流水线
```bash
python run.py
```
> 程序将自动连接 Chrome（优先端口 9223，其次 9222），嗅探当前账号并自动在 `./outputs/<账号名>_<时间戳>/` 下生成全量产物。

---

## 进阶与自定义参数

通过 `pipeline.py` 可指定端口、输出路径、条目缓存和模型精度：

```bash
python pipeline.py \
  --port 9223 \
  --output outputs/my_custom_folder \
  --model small \
  --compute-type int8
```

| 参数 | 默认值 | 说明 |
| :--- | :---: | :--- |
| `--port` | `9223` | Chrome CDP 远程调试端口 |
| `--output` | 自动生成 | 产物保存根目录 |
| `--cache` | `None` | 已收集的条目 JSON 文件路径（指定后跳过页面滚动扫描） |
| `--model` | `small` | Whisper 模型规格（`base` / `small` / `medium`） |
| `--compute-type` | `int8` | CPU 量化类型（`int8` 推理速度约 15x 实时） |

---

## 输出产物结构

每次运行完成后，目标目录将产出以下规范文件：

```text
outputs/<账号昵称>_<时间戳>/
├── douyin_favorites.xlsx     # 结构化 Excel 清单（含格式、对齐与超链接）
├── douyin_favorites.json     # 包含点赞、播放直链、音乐等全量元数据
└── transcripts/              # 逐字稿与图文 OCR 归档库
    ├── README.md             # 聚合目录索引表（相对路径可点击跳转）
    ├── 01_作者_标题.md        # 分段带时间戳逐字稿
    ├── 02_作者_标题.md        # 图文笔记多页 OCR 原文
    └── ...
```

---

## 关键技术细节与避坑指南

1. **避免平台反爬与“强制登录”拦截**：
   - 抖音官方视频详情接口对未带登录态的请求返回空数据（`强制登录`）。
   - 本工具在页面会话内部执行 `fetch(..., { credentials: 'include' })`，自动携带全部 60+ 项 Session Cookie，确保 100% 成功拉取高清流与无水印直链。
2. **PyAV 音频解码兼容性**：
   - Python 3.13 环境下 PyAV 的 `metadata_errors` 参数存在已知不兼容。
   - 本工具使用 `imageio-ffmpeg` 内存管道流拉取音轨并直接转换为 16kHz float32 NumPy 数组，零临时文件、不占磁盘空间。
3. **断点跳过**：
   - 若转录中断，再次运行同一个输出目录时，已生成且体积大于 1KB 的 Markdown 文件将自动跳过，无需重复消耗 CPU。
