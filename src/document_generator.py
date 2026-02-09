"""
文档生成器模块
"""

import re
from typing import Dict, List, Optional
from dataclasses import dataclass
import json

@dataclass
class DocumentConfig:
    """文档配置"""
    title: str
    doc_type: str
    language: str = "zh-CN"
    style: str = "academic"
    sections: Optional[List[str]] = None

class DocumentGenerator:
    """文档生成器"""
    
    def __init__(self, api_client):
        """
        初始化文档生成器
        
        Args:
            api_client: API客户端实例
        """
        self.api_client = api_client
        
        # 文档类型模板
        self.templates = {
            "期刊论文": self._journal_paper_template,
            "研究计划": self._research_proposal_template,
            "反思报告": self._reflection_template,
            "案例分析": self._case_study_template,
            "总结报告": self._summary_template
        }
    
    def generate_outline(self, 
                        title: str, 
                        doc_type: str = "期刊论文",
                        instruction: str = "") -> str:
        """
        生成文档大纲
        
        Args:
            title: 文档标题
            doc_type: 文档类型
            instruction: 附加指令
            
        Returns:
            文档大纲
        """
        # 根据文档类型选择模板
        template_func = self.templates.get(doc_type, self._journal_paper_template)
        template = template_func(title)
        
        prompt = f"""请根据以下要求生成一份{doc_type}的大纲：

论文题目：{title}

文档类型：{doc_type}

附加要求：{instruction}

请按照以下模板生成详细的大纲：
{template}

要求：
1. 大纲要详细具体，包含三级标题
2. 每个部分要说明主要内容和要点
3. 字数控制在800-1000字
4. 使用中文，格式规范

请生成完整的大纲："""
        
        system_prompt = """你是一位资深的学术编辑，擅长撰写各种学术文档的大纲。
请根据用户的要求，生成结构合理、内容详实的文档大纲。"""
        
        return self.api_client.generate_text(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=0.7,
            max_tokens=2000
        )
    
    def generate_document(self, 
                         outline: str, 
                         doc_type: str = "期刊论文") -> str:
        """
        根据大纲生成完整文档
        
        Args:
            outline: 文档大纲
            doc_type: 文档类型
            
        Returns:
            完整文档
        """
        prompt = f"""请根据以下大纲，撰写一份完整的{doc_type}：

大纲内容：
{outline}

文档类型：{doc_type}

要求：
1. 严格按照大纲结构撰写
2. 内容详实、逻辑清晰
3. 语言规范、符合学术要求
4. 包含必要的图表说明（用文字描述）
5. 字数在3000-5000字之间
6. 使用中文撰写

请开始撰写完整文档："""
        
        system_prompt = """你是一位专业的学术作家，擅长撰写高质量的学术论文、研究报告等文档。
请根据用户提供的大纲，撰写出结构完整、内容详实、语言规范的学术文档。"""
        
        return self.api_client.generate_text(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=0.8,
            max_tokens=6000
        )
    
    def _journal_paper_template(self, title: str) -> str:
        """期刊论文模板"""
        return f"""# {title}

## 摘要
- 研究背景
- 研究目的
- 研究方法
- 主要结果
- 研究结论

## 关键词
3-5个关键词

## 1. 引言
### 1.1 研究背景与意义
### 1.2 国内外研究现状
### 1.3 研究内容与目标
### 1.4 论文结构安排

## 2. 相关工作
### 2.1 相关理论研究
### 2.2 相关技术研究
### 2.3 现有方法比较
### 2.4 研究空白与创新点

## 3. 研究方法
### 3.1 研究框架
### 3.2 实验设计
### 3.3 数据收集与处理
### 3.4 分析方法

## 4. 结果与分析
### 4.1 实验结果
### 4.2 数据分析
### 4.3 结果讨论
### 4.4 假设检验

## 5. 讨论
### 5.1 结果解释
### 5.2 理论意义
### 5.3 实践意义
### 5.4 研究局限性

## 6. 结论
### 6.1 研究总结
### 6.2 主要贡献
### 6.3 未来展望

## 参考文献

## 附录（可选）"""
    
    def _research_proposal_template(self, title: str) -> str:
        """研究计划模板"""
        return f"""# {title}

## 1. 研究背景与意义
### 1.1 研究背景
### 1.2 研究意义
### 1.3 研究价值

## 2. 文献综述
### 2.1 国内外研究现状
### 2.2 研究空白
### 2.3 理论依据

## 3. 研究目标
### 3.1 总体目标
### 3.2 具体目标
### 3.3 研究问题

## 4. 研究内容与方法
### 4.1 研究内容
### 4.2 研究方法
### 4.3 技术路线
### 4.4 创新点

## 5. 研究计划
### 5.1 时间安排
### 5.2 任务分解
### 5.3 里程碑

## 6. 预期成果
### 6.1 理论成果
### 6.2 实践成果
### 6.3 社会效益

## 7. 研究基础与条件
### 7.1 研究基础
### 7.2 实验条件
### 7.3 团队组成

## 8. 参考文献

## 9. 预算（可选）"""
    
    def _reflection_template(self, title: str) -> str:
        """反思报告模板"""
        return f"""# {title}

## 1. 引言
### 1.1 反思背景
### 1.2 反思目的
### 1.3 反思意义

## 2. 事件描述
### 2.1 事件经过
### 2.2 关键环节
### 2.3 相关人员

## 3. 分析与反思
### 3.1 成功经验
### 3.2 存在问题
### 3.3 原因分析
### 3.4 改进方向

## 4. 理论联系
### 4.1 相关理论
### 4.2 理论应用
### 4.3 理论验证

## 5. 学习收获
### 5.1 知识收获
### 5.2 技能提升
### 5.3 态度转变

## 6. 行动计划
### 6.1 短期计划
### 6.2 长期计划
### 6.3 实施步骤

## 7. 结论
### 7.1 主要反思
### 7.2 未来展望

## 附录：反思日志（可选）"""
    
    def _case_study_template(self, title: str) -> str:
        """案例分析模板"""
        return f"""# {title}

## 摘要

## 1. 案例背景
### 1.1 案例来源
### 1.2 案例背景
### 1.3 研究意义

## 2. 案例描述
### 2.1 基本情况
### 2.2 发展历程
### 2.3 关键事件

## 3. 问题分析
### 3.1 主要问题
### 3.2 问题成因
### 3.3 影响分析

## 4. 理论分析
### 4.1 相关理论
### 4.2 理论应用
### 4.3 分析框架

## 5. 解决方案
### 5.1 解决方案设计
### 5.2 实施方案
### 5.3 预期效果

## 6. 讨论与启示
### 6.1 案例启示
### 6.2 推广应用
### 6.3 局限性

## 7. 结论
### 7.1 案例总结
### 7.2 实践建议

## 参考文献

## 附录：案例材料"""
    
    def _summary_template(self, title: str) -> str:
        """总结报告模板"""
        return f"""# {title}

## 1. 概述
### 1.1 总结背景
### 1.2 总结范围
### 1.3 总结目的

## 2. 工作回顾
### 2.1 主要工作
### 2.2 工作进展
### 2.3 关键成果

## 3. 成绩与经验
### 3.1 主要成绩
### 3.2 成功经验
### 3.3 亮点特色

## 4. 问题与不足
### 4.1 存在问题
### 4.2 原因分析
### 4.3 改进空间

## 5. 数据分析
### 5.1 数据概况
### 5.2 数据分析
### 5.3 数据趋势

## 6. 学习收获
### 6.1 知识学习
### 6.2 能力提升
### 6.3 认识深化

## 7. 未来展望
### 7.1 发展目标
### 7.2 工作计划
### 7.3 改进措施

## 8. 结论
### 8.1 总体评价
### 8.2 主要结论
### 8.3 建议意见

## 附录：相关数据表格"""
