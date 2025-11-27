# 修复版本检查问题 - 放在所有 import 之前
import os
import sys
import warnings

# 过滤特定警告
warnings.filterwarnings("ignore", category=UserWarning)

# 设置环境变量避免版本检查
os.environ["STREAMLIT_VERSION"] = "1.51.0"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

# 手动修复 importlib.metadata 问题
try:
    from importlib import metadata as importlib_metadata
except ImportError:
    import importlib_metadata

# 重写 version 函数以避免包元数据查找
_original_version = getattr(importlib_metadata, 'version', None)

def _patched_version(name):
    if name == "streamlit":
        return "1.51.0"
    try:
        return _original_version(name) if _original_version else "1.0.0"
    except:
        return "1.0.0"

if _original_version:
    importlib_metadata.version = _patched_version

import streamlit as st
from typing import Optional

# 导入工具模块
import tools_1  # Word+Excel批量替换工具
# 预留tools_2导入位置
# import tools_2

# 设置页面配置
st.set_page_config(
    page_title="牛马工具集",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化全局session_state
def init_global_state():
    if "active_tool" not in st.session_state:
        st.session_state.active_tool = "home"  # home, tool1, tool2

init_global_state()

# 主页面标题
st.title("📋 牛马工具集")
st.markdown("---")

# 侧边栏工具选择
with st.sidebar:
    st.header("工具选择")
    if st.button("🏠 首页", use_container_width=True):
        st.session_state.active_tool = "home"
    
    st.markdown("### 现有工具")
    if st.button("🔄 Word+Excel批量替换", use_container_width=True):
        st.session_state.active_tool = "tool1"
    
    st.markdown("### 即将上线")
    st.button("📄 wgs84-cgs2000坐标转换", use_container_width=True, disabled=True, help="敬请期待")
    # 预留tools_2入口
    # if st.button("🔧 便捷坐标转换工具", use_container_width=True):
    #     st.session_state.active_tool = "tool2"
    
    st.markdown("---")
    st.info("💡 选择左侧工具开始使用", icon="ℹ️")

# 主内容区域
if st.session_state.active_tool == "home":
    st.header("欢迎使牛马工具集")
    st.markdown("""
    本工具集提供多种牛马工作所需功能，当前已支持：
    
    - **Word+Excel批量替换**：基于Excel数据批量替换Word文档内容，支持表格内文字替换并保留格式
    
    即将推出：
    - 坐标转换工具
    
    请从左侧选择需要使用的工具开始操作。
    """)

elif st.session_state.active_tool == "tool1":
    # 调用工具1的主函数
    tools_1.main()

# 预留tools_2入口
# elif st.session_state.active_tool == "tool2":
#     tools_2.main()
