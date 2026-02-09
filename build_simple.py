#!/usr/bin/env python3
"""
简化的构建脚本，用于 GitHub Actions
"""

import sys
import os
from pathlib import Path
import subprocess
import shutil

def clean_build_dirs():
    """清理构建目录"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        dir_path = Path(dir_name)
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"Cleaned {dir_name} directory")

def build_executable():
    """构建可执行文件"""
    
    print("Starting build process...")
    
    # 确保必要的目录存在
    Path("output").mkdir(exist_ok=True)
    Path("templates").mkdir(exist_ok=True)
    Path("config").mkdir(exist_ok=True)
    
    # 清理之前的构建
    clean_build_dirs()
    
    # 构建命令
    build_args = [
        sys.executable, "-m", "PyInstaller",
        "src/main.py",
        "--name=AcademicWriterPro",
        "--onefile",
        "--windowed",
        "--add-data=src:src",
        "--hidden-import=tkinter",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=PIL",
        "--hidden-import=openai",
        "--hidden-import=requests",
        "--hidden-import=markdown",
        "--hidden-import=docx",
        "--hidden-import=tqdm",
        "--clean",
        "--noconfirm",
    ]
    
    # 根据平台调整参数
    if sys.platform == "win32":
        # Windows 特定设置
        print("Building for Windows...")
        build_args.extend([
            "--console",  # 显示控制台窗口，方便调试
        ])
    elif sys.platform == "darwin":
        # macOS 特定设置
        print("Building for macOS...")
        build_args.extend([
            "--osx-bundle-identifier=com.academicwriter.app",
        ])
    else:
        # Linux 特定设置
        print("Building for Linux...")
        build_args.extend([
            "--strip",
        ])
    
    print(f"Build command: {' '.join(build_args)}")
    
    try:
        # 运行构建命令
        result = subprocess.run(
            build_args, 
            check=True, 
            capture_output=True, 
            text=True,
            cwd=os.getcwd()
        )
        
        print("Build output:")
        print(result.stdout)
        
        if result.stderr:
            print("Build warnings/errors:")
            print(result.stderr)
        
        # 检查输出文件
        dist_dir = Path("dist")
        if dist_dir.exists():
            files = list(dist_dir.iterdir())
            print(f"\n✅ Build completed successfully!")
            print(f"Files in dist directory:")
            for file in files:
                print(f"  - {file.name} ({file.stat().st_size / 1024:.1f} KB)")
            
            # 创建压缩包
            if sys.platform == "win32":
                import zipfile
                with zipfile.ZipFile('dist/AcademicWriter-Windows.zip', 'w') as zipf:
                    for file in files:
                        zipf.write(file, arcname=file.name)
                print("Created Windows zip archive")
            
            return True
        else:
            print("❌ ERROR: dist directory not created")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed with error code: {e.returncode}")
        print(f"STDOUT: {e.stdout}")
        print(f"STDERR: {e.stderr}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print("=" * 60)
    print("Academic Writer Pro - Build Script")
    print("=" * 60)
    
    # 检查当前目录结构
    print("\nChecking project structure...")
    required_files = ["src/main.py", "requirements.txt", "src/gui.py"]
    missing_files = []
    
    for file_path in required_files:
        if not Path(file_path).exists():
            missing_files.append(file_path)
    
    if missing_files:
        print(f"❌ Missing required files: {missing_files}")
        return 1
    
    print("✅ Project structure is valid")
    
    # 开始构建
    success = build_executable()
    
    if success:
        print("\n🎉 Build completed successfully!")
        return 0
    else:
        print("\n💥 Build failed!")
        return 1

if __name__ == "__main__":
    sys.exit(main())
