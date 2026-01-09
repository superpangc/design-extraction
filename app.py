import streamlit as st
import base64
from extraction_individual import extraction_entry_stream

# 页面配置
st.set_page_config(
    page_title="PDF 解析工具",
    page_icon="📄",
    layout="centered"
)

# 自定义样式
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        padding: 0.75rem 1.5rem;
        border-radius: 10px;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
    }
    .upload-section {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
    }
    .success-box {
        background: linear-gradient(135deg, #84fab0 0%, #8fd3f4 100%);
        padding: 1.5rem;
        border-radius: 10px;
        margin-top: 1rem;
    }
    h1 {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    /* 中文化文件上传组件 */
    [data-testid="stFileUploader"] section button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.5rem 1.5rem;
        border-radius: 8px;
        font-weight: 500;
        font-size: 0;
    }
    [data-testid="stFileUploader"] section button::after {
        content: "浏览文件";
        font-size: 14px;
    }
    [data-testid="stFileUploader"] section button:hover {
        opacity: 0.9;
    }
    
    /* 隐藏所有英文提示文本 */
    [data-testid="stFileUploader"] section small {
        display: none !important;
    }
    [data-testid="stFileUploader"] section > div > div > span {
        font-size: 0 !important;
    }
    [data-testid="stFileUploader"] section > div > div > span::after {
        content: "拖拽文件到此处";
        font-size: 14px;
        color: #666;
    }
    
    /* 自定义上传区域样式 */
    [data-testid="stFileUploader"] {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px dashed #667eea;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #764ba2;
        background: #f8f9ff;
    }
    </style>
""", unsafe_allow_html=True)

# 标题和描述
st.title("📄胜利钻井设计 PDF 解析工具")
st.markdown("### 上传 PDF 文件，一键解析生成 Excel 文件")
st.markdown("---")

# 初始化 session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'excel_b64' not in st.session_state:
    st.session_state.excel_b64 = None
if 'filename' not in st.session_state:
    st.session_state.filename = None

# 文件上传区域
st.markdown('<div class="upload-section">', unsafe_allow_html=True)
st.markdown("""
    <div style='text-align: center; margin-bottom: 1rem;'>
        <p style='color: #666; font-size: 0.9rem; margin: 0;'>
            📎 拖拽文件到下方区域，或点击按钮选择文件
        </p>
        <p style='color: #999; font-size: 0.8rem; margin-top: 0.5rem;'>
            支持格式：PDF | 最大文件大小：200MB
        </p>
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "选择 PDF 文件",
    type=['pdf'],
    label_visibility="collapsed"
)
st.markdown('</div>', unsafe_allow_html=True)

# 显示上传的文件信息
if uploaded_file is not None:
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"📁 文件名: {uploaded_file.name}")
    with col2:
        file_size = len(uploaded_file.getvalue()) / 1024  # KB
        st.info(f"📊 文件大小: {file_size:.2f} KB")
    
    # 解析按钮
    if st.button("🚀 开始解析", type="primary"):
        with st.spinner("正在解析 PDF 文件，请稍候..."):
            try:
                # 读取文件的二进制内容
                pdf_binary = uploaded_file.getvalue()
                
                # 调用解析函数
                excel_b64 = extraction_entry_stream(pdf_binary)
                
                # 保存到 session state
                st.session_state.processed = True
                st.session_state.excel_b64 = excel_b64
                st.session_state.filename = uploaded_file.name.replace('.pdf', '.xlsx')
                
                st.success("✅ 解析完成！")
                
            except Exception as e:
                st.error(f"❌ 解析失败: {str(e)}")
                st.session_state.processed = False

# 显示下载链接
if st.session_state.processed and st.session_state.excel_b64:
    st.markdown("---")
    st.markdown('<div class="success-box">', unsafe_allow_html=True)
    st.markdown("### 🎉 解析成功！")
    st.markdown(f"**生成的文件:** {st.session_state.filename}")
    
    # 解码 base64 为二进制
    excel_binary = base64.b64decode(st.session_state.excel_b64)
    
    # 创建下载按钮
    st.download_button(
        label="📥 下载 Excel 文件",
        data=excel_binary,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# 页脚
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; padding: 1rem;'>
        <p>💡 提示：支持上传 PDF 文件，解析后生成 Excel 格式的报告</p>
    </div>
    """,
    unsafe_allow_html=True
)
