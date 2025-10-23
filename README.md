# Subtitle-Exporter
快速批量导出字幕

---

## 环境准备

- **安装ffmpeg**:
  - 官网：https://ffmpeg.org/?utm_source=chatgpt.com
  - 安装到Path(建议)：下载并解压到任意目录后将包含ffmpeg.exe和ffprobe.exe的目录添加到Path
  - 安装到程序目录：下载并解压ffmpeg.exe和ffprobe.exe到本工具exe文件所在目录下的ffmpeg文件夹内(确保ffmpeg.exe和ffprobe.exe位于./ffmpeg/ffmpeg.exe和./ffmpeg/ffprobe.exe)
- **安装spp2pgs**:
  - 官网：https://github.com/subelf/Spp2Pgs/releases
  - 安装到Path(建议)：下载并解压到任意目录后将包含spp2pgs.exe的目录添加到Path
  - 安装到程序目录：载并解压整个目录到本工具exe文件所在目录下的spp2pgs文件夹内(确保spp2pgs.exe位于./spp2pgs/spp2pgs.exe)
- **安装FontForge**:
  - 官网：https://fontforge.org/en-US/?utm_source=chatgpt.com
  - 安装到Path:下载安装包并安装，将安装目录下的bin目录添加到Path
  - 安装到程序目录(建议)：下载安装包并安装，将整个安装目录复制到本工具exe文件所在目录下的FontForge文件夹内(确保fontforge.exe位于./FontForge/bin/fontforge.exe)，完成后为FontForge文件夹添加Everyone读写权限

---

## 使用程序
<img width="902" height="648" alt="屏幕截图 2025-10-24 004507" src="https://github.com/user-attachments/assets/1d59ab9b-3f3b-4812-913e-ee5e50717725" />

- **导入文件**:
  - 使用左上角**导入文件**按钮打开文件选择窗口导入文件（支持批量导入）
  - 直接向窗口**拖入**文件（支持多选拖入）
- **编辑信息**:
  - 单击文件选中行（高亮的行）
  - 再次单机选中行的“**宽度**”，“**高度**”，“**刷新率**”列可对其中参数进行修改
  - 按住**shift**或**ctrl**可多选行
  - 点击左上角**删除**按钮可删除选中的文件（高亮的行），点击**清空**按钮可清空全部文件
  - 点击文件第一列的**复选框**可**选中/取消选中**文件（此处选中的文件为要导出字幕的文件）
  - 点击表头的**复选框**可**全选/全不选**文件
