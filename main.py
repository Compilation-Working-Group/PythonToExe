import customtkinter as ctk
import pubchempy as pcp
from rdkit import Chem
from rdkit.Chem import AllChem
import webbrowser
import os
from deep_translator import GoogleTranslator
import threading
import sys

# --- 配置区域 (已修正开发者信息) ---
APP_VERSION = "v1.1.0"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室" 
DEV_SCHOOL = "金塔县中学"
COPYRIGHT_YEAR = "2026"
# ---------------------------------------------

# 设置外观模式
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AboutWindow(ctk.CTkToplevel):
    """关于软件的弹窗"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("关于软件")
        self.geometry("400x300")
        self.resizable(False, False)
        
        # 保持窗口在最前
        self.attributes("-topmost", True)

        # 标题
        self.label_title = ctk.CTkLabel(self, text="有机分子结构 3D 展示工具", font=("Microsoft YaHei UI", 18, "bold"))
        self.label_title.pack(pady=(20, 10))
        
        self.label_ver = ctk.CTkLabel(self, text=f"版本: {APP_VERSION}", font=("Arial", 12))
        self.label_ver.pack(pady=0)

        # 分割线
        self.frame_line = ctk.CTkFrame(self, height=2, fg_color="gray")
        self.frame_line.pack(fill="x", padx=50, pady=15)

        # 开发者信息
        info_text = f"开发者: {DEV_NAME}\n单位: {DEV_SCHOOL}\n{DEV_ORG}"
        self.label_dev = ctk.CTkLabel(self, text=info_text, font=("Microsoft YaHei UI", 14), justify="center")
        self.label_dev.pack(pady=10)

        # 技术致谢
        credits_text = "Powered by: Python, RDKit, PubChemPy, 3Dmol.js\n自动化构建: GitHub Actions"
        self.label_credits = ctk.CTkLabel(self, text=credits_text, font=("Arial", 10), text_color="gray")
        self.label_credits.pack(side="bottom", pady=20)

class MoleculeViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"有机分子 3D 教学助手 - {DEV_NAME}作品")
        self.geometry("640x500")
        self.grid_columnconfigure(0, weight=1)
        self.toplevel_window = None

        # --- 顶部布局 ---
        self.frame_top = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_top.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        # 标题 (左对齐)
        self.label_title = ctk.CTkLabel(self.frame_top, text="有机化学分子 3D 建模", font=("Microsoft YaHei UI", 24, "bold"))
        self.label_title.pack(side="left")

        # 关于按钮 (右对齐)
        self.btn_about = ctk.CTkButton(self.frame_top, text="关于 / About", width=80, height=24, 
                                       fg_color="transparent", border_width=1, 
                                       text_color=("gray10", "gray90"), command=self.open_about)
        self.btn_about.pack(side="right")

        # --- 主体内容 ---
        # 说明
        self.label_desc = ctk.CTkLabel(self, text="输入中文名称、英文名称、分子式或 SMILES\n(例如: 苯酚, 乙酸乙酯, Aspirin)", 
                                       font=("Microsoft YaHei UI", 14), text_color="gray")
        self.label_desc.grid(row=1, column=0, padx=20, pady=10)

        # 输入框
        self.entry_chem = ctk.CTkEntry(self, placeholder_text="在此输入有机物名称...", width=450, height=45, font=("Microsoft YaHei UI", 16))
        self.entry_chem.grid(row=2, column=0, padx=20, pady=15)
        self.entry_chem.bind("<Return>", self.start_generation_thread)

        # 样式选择框
        self.style_frame = ctk.CTkFrame(self)
        self.style_frame.grid(row=3, column=0, pady=10)
        self.label_style = ctk.CTkLabel(self.style_frame, text="模型样式:", font=("Microsoft YaHei UI", 12, "bold"))
        self.label_style.pack(side="left", padx=10)
        
        self.style_var = ctk.StringVar(value="stick")
        ctk.CTkRadioButton(self.style_frame, text="球棍模型 (键角明显)", variable=self.style_var, value="stick").pack(side="left", padx=10, pady=10)
        ctk.CTkRadioButton(self.style_frame, text="比例模型 (体积明显)", variable=self.style_var, value="sphere").pack(side="left", padx=10, pady=10)

        # 生成按钮
        self.btn_generate = ctk.CTkButton(self, text="立即生成 3D 结构", command=self.start_generation_thread, 
                                          height=50, width=200, font=("Microsoft YaHei UI", 18, "bold"))
        self.btn_generate.grid(row=4, column=0, padx=20, pady=20)

        # 状态栏
        self.status_label = ctk.CTkLabel(self, text="准备就绪", text_color="gray")
        self.status_label.grid(row=5, column=0, pady=10)
        
        # 底部版权
        self.label_footer = ctk.CTkLabel(self, text=f"© {COPYRIGHT_YEAR} {DEV_ORG}", font=("Microsoft YaHei UI", 10), text_color="gray50")
        self.label_footer.grid(row=6, column=0, pady=(20, 10))

    def open_about(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = AboutWindow(self)
        else:
            self.toplevel_window.focus()

    def start_generation_thread(self, event=None):
        threading.Thread(target=self.generate_model, daemon=True).start()

    def generate_model(self):
        user_input = self.entry_chem.get().strip()
        if not user_input:
            self.status_label.configure(text="请输入内容！", text_color="red")
            return

        self.status_label.configure(text=f"正在搜索 '{user_input}' ...", text_color="#1F6AA5")
        self.btn_generate.configure(state="disabled", text="正在计算...")

        try:
            # 1. 翻译
            search_query = user_input
            if self.is_contains_chinese(user_input):
                try:
                    search_query = GoogleTranslator(source='auto', target='en').translate(user_input)
                except Exception:
                    pass 
            
            # 2. 搜索
            compounds = pcp.get_compounds(search_query, 'name')
            if not compounds:
                compounds = pcp.get_compounds(search_query, 'formula')
            
            if not compounds:
                self.status_label.configure(text=f"未找到 '{user_input}'，请尝试输入英文或分子式。", text_color="orange")
                self.btn_generate.configure(state="normal", text="立即生成 3D 结构")
                return

            target_compound = compounds[0]
            smiles = target_compound.canonical_smiles
            name = user_input 

            # 3. RDKit 处理
            mol = Chem.MolFromSmiles(smiles)
            if mol is None:
                raise ValueError("无法解析分子结构")

            mol_with_h = Chem.AddHs(mol)
            
            res = AllChem.EmbedMolecule(mol_with_h, AllChem.ETKDG())
            if res == -1:
                AllChem.EmbedMolecule(mol_with_h, AllChem.ETKDG(), useRandomCoords=True)

            mol_block = Chem.MolToMolBlock(mol_with_h)

            # 4. 生成 HTML
            self.create_html_viewer(name, mol_block, self.style_var.get())
            
            self.status_label.configure(text=f"成功！已打开 {name}", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
            print(e)
        finally:
            self.btn_generate.configure(state="normal", text="立即生成 3D 结构")

    def is_contains_chinese(self, strs):
        for _char in strs:
            if '\u4e00' <= _char <= '\u9fa5':
                return True
        return False

    def create_html_viewer(self, title, mol_data, style):
        # 针对 MOL 格式调整样式配置
        style_config = ""
        if style == "stick":
            style_config = "viewer.setStyle({}, {stick: {radius: 0.14, colorscheme: 'Jmol'}, sphere: {scale: 0.23, colorscheme: 'Jmol'}});"
        else:
            style_config = "viewer.setStyle({}, {sphere: {colorscheme: 'Jmol'}});"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>{title} - {DEV_NAME} 3D 演示</title>
            <script src="https://3Dmol.org/build/3Dmol-min.js"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; background-color: #f5f7fa; font-family: "Microsoft YaHei", sans-serif; }}
                #container {{ width: 100vw; height: 100vh; position: relative; }}
                #info {{ 
                    position: absolute; top: 20px; left: 20px; z-index: 10; 
                    background: rgba(255, 255, 255, 0.95); padding: 15px 20px; 
                    border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); 
                    border-left: 5px solid #3B8ED0;
                }}
                h2 {{ margin: 0 0 5px 0; color: #2c3e50; font-size: 22px; }}
                p {{ margin: 5px 0; font-size: 14px; color: #7f8c8d; }}
                .legend {{ margin-top: 15px; font-size: 13px; display: flex; gap: 10px; }}
                .legend-item {{ display: flex; align-items: center; }}
                .dot {{ height: 12px; width: 12px; display: inline-block; border-radius: 50%; margin-right: 6px; border: 1px solid rgba(0,0,0,0.1); }}
                .footer {{ margin-top: 10px; font-size: 12px; color: #bdc3c7; text-align: right; border-top: 1px solid #eee; padding-top: 5px;}}
            </style>
        </head>
        <body>
            <div id="info">
                <h2>{title}</h2>
                <p>🖱️ 左键旋转 | 🖱️ 滚轮缩放 | 🖱️ 右键平移</p>
                <div class="legend">
                    <div class="legend-item"><span class="dot" style="background:#909090;"></span>C</div>
                    <div class="legend-item"><span class="dot" style="background:#FFFFFF;"></span>H</div>
                    <div class="legend-item"><span class="dot" style="background:#FF0D0D;"></span>O</div>
                    <div class="legend-item"><span class="dot" style="background:#3050F8;"></span>N</div>
                </div>
                <div class="footer">Design by {DEV_ORG}</div>
            </div>
            <div id="container" class="mol-container"></div>
            <script>
                let element = document.getElementById('container');
                let config = {{ backgroundColor: '#f5f7fa' }};
                let viewer = $3Dmol.createViewer(element, config);
                let molData = `{mol_data}`;
                viewer.addModel(molData, "mol");
                {style_config}
                viewer.zoomTo();
                viewer.render();
            </script>
        </body>
        </html>
        """
        
        filename = "structure_view.html"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        webbrowser.open('file://' + os.path.realpath(filename))

if __name__ == "__main__":
    app = MoleculeViewerApp()
    app.mainloop()
