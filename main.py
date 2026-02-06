import customtkinter as ctk
import pubchempy as pcp
from rdkit import Chem
from rdkit.Chem import AllChem
import webbrowser
import os
from deep_translator import GoogleTranslator
import threading

# 设置外观模式
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class MoleculeViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("有机分子结构 3D 展示工具 - 教学版")
        self.geometry("600x450")
        self.grid_columnconfigure(0, weight=1)

        # 标题
        self.label_title = ctk.CTkLabel(self, text="有机化学分子 3D 建模助手", font=("Microsoft YaHei UI", 24, "bold"))
        self.label_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 说明
        self.label_desc = ctk.CTkLabel(self, text="输入中文名称、英文名称、分子式或 SMILES (例如: 甲烷, 乙醇)", font=("Microsoft YaHei UI", 14))
        self.label_desc.grid(row=1, column=0, padx=20, pady=10)

        # 输入框
        self.entry_chem = ctk.CTkEntry(self, placeholder_text="在此输入有机物名称...", width=400, height=40)
        self.entry_chem.grid(row=2, column=0, padx=20, pady=10)
        self.entry_chem.bind("<Return>", self.start_generation_thread)

        # 样式选择
        self.style_var = ctk.StringVar(value="stick")
        self.radio_frame = ctk.CTkFrame(self)
        self.radio_frame.grid(row=3, column=0, pady=10)
        
        ctk.CTkRadioButton(self.radio_frame, text="球棍模型 (Ball & Stick)", variable=self.style_var, value="stick").pack(side="left", padx=10)
        ctk.CTkRadioButton(self.radio_frame, text="空间填充 (Sphere)", variable=self.style_var, value="sphere").pack(side="left", padx=10)

        # 按钮
        self.btn_generate = ctk.CTkButton(self, text="生成并查看 3D 模型", command=self.start_generation_thread, height=40, font=("Microsoft YaHei UI", 16, "bold"))
        self.btn_generate.grid(row=4, column=0, padx=20, pady=20)

        # 状态显示
        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=5, column=0, pady=10)

    def start_generation_thread(self, event=None):
        threading.Thread(target=self.generate_model, daemon=True).start()

    def generate_model(self):
        user_input = self.entry_chem.get().strip()
        if not user_input:
            self.status_label.configure(text="请输入内容！", text_color="red")
            return

        self.status_label.configure(text=f"正在搜索 '{user_input}' ...", text_color="orange")
        self.btn_generate.configure(state="disabled")

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
                self.status_label.configure(text="未找到该化合物，请检查拼写。", text_color="red")
                self.btn_generate.configure(state="normal")
                return

            target_compound = compounds[0]
            smiles = target_compound.canonical_smiles
            name = user_input 

            # 3. RDKit 处理
            mol = Chem.MolFromSmiles(smiles)
            if mol is None:
                raise ValueError("无法解析分子结构")

            # 关键步骤：添加氢原子 (AddHs)
            mol_with_h = Chem.AddHs(mol)
            
            # 计算 3D 坐标
            # 使用 ETKDG 算法生成构象
            res = AllChem.EmbedMolecule(mol_with_h, AllChem.ETKDG())
            if res == -1:
                # 如果标准算法失败，尝试随机坐标（针对复杂分子）
                AllChem.EmbedMolecule(mol_with_h, AllChem.ETKDG(), useRandomCoords=True)

            # 关键修改：导出为 MOL Block 格式，而不是 PDB
            # MOL 格式对小分子支持更好，显式包含所有原子
            mol_block = Chem.MolToMolBlock(mol_with_h)

            # 4. 生成 HTML
            self.create_html_viewer(name, mol_block, self.style_var.get())
            
            self.status_label.configure(text=f"成功！已在浏览器中打开 {name}", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
            print(e)
        finally:
            self.btn_generate.configure(state="normal")

    def is_contains_chinese(self, strs):
        for _char in strs:
            if '\u4e00' <= _char <= '\u9fa5':
                return True
        return False

    def create_html_viewer(self, title, mol_data, style):
        # 针对 MOL 格式调整样式配置
        style_config = ""
        if style == "stick":
            # 这种配置下，MOL 格式会正确渲染 C 和 H
            style_config = "viewer.setStyle({}, {stick: {radius: 0.15, colorscheme: 'Jmol'}, sphere: {scale: 0.25, colorscheme: 'Jmol'}});"
        else:
            style_config = "viewer.setStyle({}, {sphere: {colorscheme: 'Jmol'}});"

        # 使用 JavaScript 的反引号 (template literals) 也可以处理多行字符串，
        # 但为了安全起见，我们将 Python 字符串直接传给 JS 变量
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>{title} - 3D Structure</title>
            <script src="https://3Dmol.org/build/3Dmol-min.js"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; background-color: #f0f2f5; font-family: "Microsoft YaHei", sans-serif; }}
                #container {{ width: 100vw; height: 100vh; position: relative; }}
                #info {{ 
                    position: absolute; top: 20px; left: 20px; z-index: 10; 
                    background: rgba(255, 255, 255, 0.9); padding: 15px; 
                    border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
                }}
                h2 {{ margin: 0 0 5px 0; color: #333; }}
                p {{ margin: 0; font-size: 14px; color: #666; }}
                .legend {{ margin-top: 10px; font-size: 12px; }}
                .dot {{ height: 10px; width: 10px; display: inline-block; border-radius: 50%; margin-right: 5px; }}
            </style>
        </head>
        <body>
            <div id="info">
                <h2>{title}</h2>
                <p>左键旋转 | 滚轮缩放 | 右键平移</p>
                <div class="legend">
                    <div><span class="dot" style="background:#909090;"></span>碳 (C)</div>
                    <div><span class="dot" style="background:#FFFFFF; border:1px solid #ccc;"></span>氢 (H)</div>
                    <div><span class="dot" style="background:#FF0D0D;"></span>氧 (O)</div>
                    <div><span class="dot" style="background:#3050F8;"></span>氮 (N)</div>
                </div>
            </div>
            <div id="container" class="mol-container"></div>
            <script>
                let element = document.getElementById('container');
                let config = {{ backgroundColor: 'white' }};
                let viewer = $3Dmol.createViewer(element, config);
                
                // 使用 MOL 格式数据
                let molData = `{mol_data}`;
                
                // 注意这里改成了 "mol"
                viewer.addModel(molData, "mol");
                
                {style_config}
                
                viewer.zoomTo();
                viewer.render();
            </script>
        </body>
        </html>
        """
        
        filename = "structure_view.html"
        # 确保使用 utf-8 写入
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        webbrowser.open('file://' + os.path.realpath(filename))

if __name__ == "__main__":
    app = MoleculeViewerApp()
    app.mainloop()
