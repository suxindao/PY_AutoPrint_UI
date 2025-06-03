### 打包步骤
    确保所有依赖项已安装：
    
    bash
    pip install pywin32 pyinstaller pythoncom
    在项目根目录运行打包命令：
    
    bash
    pyinstaller --onefile --windowed --icon=icon.ico --add-data "config;config" --add-data "logs;logs" main.py
    或者使用spec文件：
    
    bash
    pyinstaller build.spec
    打包完成后，EXE文件会生成在dist文件夹中

### 高级配置选项
    添加程序图标
    准备一个.ico格式的图标文件（如icon.ico）
    
    在打包命令中添加--icon=icon.ico参数
    
    隐藏控制台窗口
    在spec文件中设置console=False或在命令行添加--windowed参数
    
    包含数据文件
    确保配置文件和日志目录被包含：
    
    bash
    --add-data "config;config" --add-data "logs;logs"

### 从左到右向日葵

    688 015 466 (外星人)
    910 380 437 (英文)
    657 098 865 (何总)
    321 567 317 (高速打印机)
