"""
DeepSeek API客户端模块
"""

import os
import json
from typing import Dict, List, Optional, Any
from openai import OpenAI
import requests

class DeepSeekClient:
    """DeepSeek API客户端"""
    
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com"):
        """
        初始化DeepSeek客户端
        
        Args:
            api_key: DeepSeek API密钥
            base_url: API基础URL
        """
        self.api_key = api_key
        self.base_url = base_url
        self.client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )
    
    def chat_completion(self, 
                       messages: List[Dict[str, str]],
                       model: str = "deepseek-chat",
                       temperature: float = 0.7,
                       max_tokens: int = 4000,
                       stream: bool = False) -> Dict[str, Any]:
        """
        聊天补全接口
        
        Args:
            messages: 消息列表
            model: 模型名称
            temperature: 温度参数
            max_tokens: 最大token数
            stream: 是否流式输出
            
        Returns:
            API响应结果
        """
        try:
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                stream=stream
            )
            
            if stream:
                return response
            
            return {
                "content": response.choices[0].message.content,
                "usage": response.usage.dict() if response.usage else None
            }
            
        except Exception as e:
            raise Exception(f"API调用失败: {str(e)}")
    
    def generate_text(self, 
                     prompt: str,
                     system_prompt: str = "你是一个专业的学术写作助手。",
                     model: str = "deepseek-chat",
                     temperature: float = 0.7,
                     max_tokens: int = 4000) -> str:
        """
        生成文本
        
        Args:
            prompt: 用户提示
            system_prompt: 系统提示
            model: 模型名称
            temperature: 温度参数
            max_tokens: 最大token数
            
        Returns:
            生成的文本
        """
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ]
        
        response = self.chat_completion(
            messages=messages,
            model=model,
            temperature=temperature,
            max_tokens=max_tokens
        )
        
        return response["content"]
    
    def check_api_key(self) -> bool:
        """
        检查API密钥是否有效
        
        Returns:
            bool: API密钥是否有效
        """
        try:
            # 发送一个简单的测试请求
            test_prompt = "Hello"
            response = self.generate_text(
                prompt=test_prompt,
                max_tokens=10
            )
            return response is not None
        except:
            return False
