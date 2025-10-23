# Subtitle-Exporter
快速批量导出字幕

---

## 环境准备

- **安装ffmpeg**:
  - 官网：https://ffmpeg.org/?utm_source=chatgpt.com
  - 安装到Path(建议)：下载并解压到任意目录后将包含ffmpeg.exe和ffprobe.exe的目录添加到Path
  - 安装到程序目录：下载并解压ffmpeg.exe和ffprobe.exe到本工具exe文件所在目录下的ffmpeg文件夹内
- **安装spp2pgs**:
  - 官网：https://github.com/subelf/Spp2Pgs/releases
  - 安装到Path(建议)：下载并解压到任意目录后将包含spp2pgs.exe的目录添加到Path
  - 安装到程序目录：载并解压整个目录到本工具exe文件所在目录下的spp2pgs文件夹内(确保spp2pgs.exe位于./spp2pgs/spp2pgs.exe)
- **安装FontForge**:
  - 官网：https://fontforge.org/en-US/?utm_source=chatgpt.com
  - 安装到Path:下载安装包并安装，将安装目录下的bin目录添加到Path
  - 安装到程序目录(建议)：下载安装包并安装，将整个安装目录复制到本工具exe文件所在目录下的FontForge文件夹内(确保fontforge.exe位于./FontForge/bin/fontforge.exe)，完成后为FontForge文件夹添加Everyone读写权限
