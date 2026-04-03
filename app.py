"""
语文练习册生成器 — Streamlit 图形界面 v2
运行：streamlit run app.py
"""

import anthropic
import streamlit as st
from builder import (
    build_document,
    SUBJECT_READING, SUBJECT_WRITING, SUBJECT_BASIC,
    SUBJECT_CLASSICS, SUBJECT_WENYAN, SUBJECT_POETRY,
)

# ── 页面配置 ──────────────────────────────────────────────────
st.set_page_config(
    page_title='语文练习册生成器',
    page_icon='📚',
    layout='wide',
)

ALL_SUBJECTS = [
    SUBJECT_READING, SUBJECT_CLASSICS, SUBJECT_WENYAN, SUBJECT_POETRY,
    SUBJECT_WRITING, SUBJECT_BASIC,
]


# ══════════════════════════════════════════════════════════════
# 数据初始化工具
# ══════════════════════════════════════════════════════════════

def _new_lecture(number):
    return {
        'number':            number,
        'subject':           SUBJECT_READING,
        'topic':             '',
        'core_method_lines': 8,
        'sections': [_new_section('巩固提升'), _new_section('能力进阶')],
        'daily_words':       [],
    }


def _new_section(name):
    return {
        'name':            name,
        'time_suggestion': '',
        'articles':        [_new_article()],
        'questions':       [_new_question()],
        'prompts':         [_new_prompt()],
    }


def _new_article():
    return {'title': '', 'author': '', 'body': '', 'source': '', 'image': None}


def _new_question():
    return {'type': '核心题', 'text': '', 'linked_material': '',
            'hint': '', 'answer_lines': 4, 'image': None, 'table': None, 'answer': ''}


def _new_prompt():
    return {'text': '', 'requirements': '', 'writing_lines': 22, 'sample_essay': ''}


def _generate_hint(articles, question_text):
    """调用 Claude API，根据文章和题目自动生成约 50 字答题提示。"""
    article_text = '\n'.join(
        art.get('body', '') for art in articles if art.get('body', '').strip()
    )
    client = anthropic.Anthropic()
    msg = client.messages.create(
        model='claude-haiku-4-5-20251001',
        max_tokens=150,
        messages=[{
            'role': 'user',
            'content': (
                f'以下是阅读文章：\n{article_text}\n\n'
                f'题目：{question_text}\n\n'
                '请为这道题生成一条答题提示，帮助学生找到答题角度和方法，'
                '约50字，不要直接给出答案，只提供思路方向。直接输出提示内容，无需任何前缀。'
            ),
        }],
    )
    return msg.content[0].text.strip()


def _new_word():
    return {'pinyin': '', 'hanzi': ''}


# ══════════════════════════════════════════════════════════════
# UI 渲染函数（先定义，后调用）
# ══════════════════════════════════════════════════════════════

def render_article_question_tabs(lec_i, lec):
    """渲染阅读类板块（文章 + 题目）"""
    section_tabs = st.tabs([f"📖 {sec['name']}" for sec in lec['sections'][:2]])
    for sec_i, (tab, sec) in enumerate(zip(section_tabs, lec['sections'][:2])):
        with tab:
            s_c1, s_c2 = st.columns([2, 1])
            with s_c1:
                sec['name'] = st.text_input('部分名称', value=sec['name'],
                                            key=f's_name_{lec_i}_{sec_i}')
            with s_c2:
                sec['time_suggestion'] = st.text_input(
                    '建议用时', value=sec.get('time_suggestion', ''),
                    placeholder='如：15分钟', key=f's_time_{lec_i}_{sec_i}')

            # ── 文章 ──
            st.markdown('##### 阅读文章')
            for art_i, art in enumerate(sec.get('articles', [])):
                with st.container(border=True):
                    ac1, ac2, ac3 = st.columns([2, 1, 1])
                    with ac1:
                        art['title'] = st.text_input(
                            '文章标题', value=art.get('title', ''),
                            key=f'art_t_{lec_i}_{sec_i}_{art_i}')
                    with ac2:
                        art['author'] = st.text_input(
                            '作者', value=art.get('author', ''),
                            placeholder='如：鲁迅',
                            key=f'art_a_{lec_i}_{sec_i}_{art_i}')
                    with ac3:
                        art['source'] = st.text_input(
                            '出处', value=art.get('source', ''),
                            placeholder='（选自《...》，有删改）',
                            key=f'art_s_{lec_i}_{sec_i}_{art_i}')
                    art['body'] = st.text_area(
                        '文章正文（每段一行）',
                        value=art.get('body', ''), height=200,
                        key=f'art_b_{lec_i}_{sec_i}_{art_i}')
                    st.caption('格式：`__文字__` → 下划线　｜　独立行 `---` → 横线')
                    art_img = st.file_uploader(
                        '文章配图（可选）', type=['png', 'jpg', 'jpeg'],
                        key=f'art_img_{lec_i}_{sec_i}_{art_i}')
                    art['image'] = art_img.read() if art_img else None

                    if len(sec['articles']) > 1:
                        if st.button('🗑 删除文章',
                                     key=f'del_art_{lec_i}_{sec_i}_{art_i}'):
                            sec['articles'].pop(art_i)
                            st.rerun()

            if st.button('➕ 添加文章', key=f'add_art_{lec_i}_{sec_i}'):
                sec['articles'].append(_new_article())
                st.rerun()

            st.markdown('---')

            # ── 题目 ──
            st.markdown('##### 练习题目')
            for q_i, q in enumerate(sec.get('questions', [])):
                with st.container(border=True):
                    qc1, qc2, qc3 = st.columns([1, 3, 1])
                    with qc1:
                        q['type'] = st.selectbox(
                            '题型', ['核心题', '热搜题', '新趋势'],
                            index=['核心题', '热搜题', '新趋势'].index(
                                q.get('type', '核心题')),
                            key=f'qt_{lec_i}_{sec_i}_{q_i}')
                    with qc2:
                        q['text'] = st.text_area(
                            f'第 {q_i+1} 题题目',
                            value=q.get('text', ''), height=80,
                            key=f'qtxt_{lec_i}_{sec_i}_{q_i}')
                        st.caption('`__文字__` → 下划线')
                    with qc3:
                        q['answer_lines'] = st.number_input(
                            '答题行数', min_value=1, max_value=20,
                            value=q.get('answer_lines', 4),
                            key=f'qlines_{lec_i}_{sec_i}_{q_i}')

                    # ── 表格编辑器 ──
                    use_table = st.checkbox(
                        '插入表格', value=q.get('table') is not None,
                        key=f'use_table_{lec_i}_{sec_i}_{q_i}')
                    if use_table:
                        if q.get('table') is None:
                            q['table'] = {'has_header': True,
                                          'data': [['', ''], ['', '']]}
                        td = q['table']
                        tc1, tc2, tc3 = st.columns([1, 1, 3])
                        with tc1:
                            new_rows = int(st.number_input(
                                '行数', min_value=1, max_value=10,
                                value=max(1, len(td['data'])),
                                key=f'trows_{lec_i}_{sec_i}_{q_i}'))
                        with tc2:
                            cur_cols = len(td['data'][0]) if td['data'] else 2
                            new_cols = int(st.number_input(
                                '列数', min_value=1, max_value=8,
                                value=max(1, cur_cols),
                                key=f'tcols_{lec_i}_{sec_i}_{q_i}'))
                        with tc3:
                            td['has_header'] = st.checkbox(
                                '首行为表头', value=td.get('has_header', True),
                                key=f'thead_{lec_i}_{sec_i}_{q_i}')
                        # 自动调整行列数
                        while len(td['data']) < new_rows:
                            td['data'].append(
                                [''] * (len(td['data'][0]) if td['data'] else new_cols))
                        td['data'] = td['data'][:new_rows]
                        for row in td['data']:
                            while len(row) < new_cols:
                                row.append('')
                            del row[new_cols:]
                        # 单元格输入网格
                        for r_i, row in enumerate(td['data']):
                            is_hdr = td.get('has_header') and r_i == 0
                            cell_cols = st.columns(len(row))
                            for c_i, cell_val in enumerate(row):
                                with cell_cols[c_i]:
                                    td['data'][r_i][c_i] = st.text_input(
                                        f'{"表头" if is_hdr else "内容"}({r_i+1},{c_i+1})',
                                        value=cell_val,
                                        key=f'tcell_{lec_i}_{sec_i}_{q_i}_{r_i}_{c_i}',
                                        label_visibility='collapsed')
                    else:
                        q['table'] = None

                    # ── 题目配图 ──
                    q_img = st.file_uploader(
                        '题目配图（可选）', type=['png', 'jpg', 'jpeg'],
                        key=f'qimg_{lec_i}_{sec_i}_{q_i}')
                    q['image'] = q_img.read() if q_img else None

                    # ── 链接材料 ──
                    q['linked_material'] = st.text_area(
                        '链接材料（可留空）', value=q.get('linked_material', ''),
                        height=60, key=f'qlink_{lec_i}_{sec_i}_{q_i}')

                    # ── 答题提示 ──
                    hint_col, btn_col = st.columns([5, 1])
                    with hint_col:
                        q['hint'] = st.text_area(
                            '答题提示（AI 自动生成）',
                            value=q.get('hint', ''),
                            height=68,
                            key=f'qhint_{lec_i}_{sec_i}_{q_i}')
                    with btn_col:
                        st.markdown('<div style="margin-top:28px"></div>', unsafe_allow_html=True)
                        if st.button('✨ 生成', key=f'gen_hint_{lec_i}_{sec_i}_{q_i}'):
                            _err = None
                            with st.spinner('生成中…'):
                                try:
                                    generated = _generate_hint(
                                        sec.get('articles', []),
                                        q.get('text', ''))
                                    q['hint'] = generated
                                    st.session_state[f'qhint_{lec_i}_{sec_i}_{q_i}'] = generated
                                except Exception as e:
                                    _err = e
                            if _err:
                                st.error(f'生成失败：{_err}')
                            else:
                                st.rerun()

                    # ── 参考答案 ──
                    q['answer'] = st.text_area(
                        '参考答案（将汇总排在文档末尾）',
                        value=q.get('answer', ''),
                        height=68,
                        key=f'qans_{lec_i}_{sec_i}_{q_i}')

                    if len(sec['questions']) > 1:
                        if st.button('🗑 删除此题',
                                     key=f'del_q_{lec_i}_{sec_i}_{q_i}'):
                            sec['questions'].pop(q_i)
                            st.rerun()

            if st.button('➕ 添加题目', key=f'add_q_{lec_i}_{sec_i}'):
                sec['questions'].append(_new_question())
                st.rerun()


def render_writing_tabs(lec_i, lec):
    """渲染作文板块（写作题目）"""
    section_tabs = st.tabs(['📖 巩固提升', '📖 能力进阶'])
    for sec_i, (tab, sec) in enumerate(zip(section_tabs, lec['sections'][:2])):
        with tab:
            s_c1, s_c2 = st.columns([2, 1])
            with s_c1:
                sec['name'] = st.text_input('部分名称', value=sec['name'],
                                            key=f's_name_{lec_i}_{sec_i}')
            with s_c2:
                sec['time_suggestion'] = st.text_input(
                    '建议用时', value=sec.get('time_suggestion', ''),
                    placeholder='如：15分钟', key=f's_time_{lec_i}_{sec_i}')

            st.markdown('##### 写作题目')
            for p_i, prompt in enumerate(sec.get('prompts', [])):
                with st.container(border=True):
                    prompt['text'] = st.text_area(
                        '写作题目', value=prompt.get('text', ''),
                        height=80, key=f'prompt_{lec_i}_{sec_i}_{p_i}')
                    st.caption('格式：`__文字__` → 下划线　｜　独立行 `---` → 横线')
                    prompt['requirements'] = st.text_area(
                        '写作要求（可选）', value=prompt.get('requirements', ''),
                        height=60, key=f'req_{lec_i}_{sec_i}_{p_i}')
                    prompt['writing_lines'] = st.slider(
                        '空白书写行数', 10, 40,
                        prompt.get('writing_lines', 22),
                        key=f'wlines_{lec_i}_{sec_i}_{p_i}')
                    prompt['sample_essay'] = st.text_area(
                        '参考范文（将汇总排在文档末尾，可留空）',
                        value=prompt.get('sample_essay', ''),
                        height=150,
                        key=f'essay_{lec_i}_{sec_i}_{p_i}')
                    st.caption('格式：`__文字__` → 下划线　｜　独立行 `---` → 横线')

                    if len(sec['prompts']) > 1:
                        if st.button('🗑 删除题目',
                                     key=f'del_prompt_{lec_i}_{sec_i}_{p_i}'):
                            sec['prompts'].pop(p_i)
                            st.rerun()

            if st.button('➕ 添加写作题目', key=f'add_prompt_{lec_i}_{sec_i}'):
                sec['prompts'].append(_new_prompt())
                st.rerun()


def render_daily_words(lec_i, lec):
    """渲染日积月累词语输入"""
    st.markdown('##### 📖 日积月累 — 字词描红')
    st.caption('填写拼音和汉字，将生成灰色描红表格供学生临摹，下方留空格独立默写')

    header_c1, header_c2, _ = st.columns([2, 2, 1])
    with header_c1:
        st.markdown('**拼音**')
    with header_c2:
        st.markdown('**汉字**')

    words = lec.setdefault('daily_words', [])
    for w_i, word in enumerate(words):
        wc1, wc2, wc3 = st.columns([2, 2, 1])
        with wc1:
            word['pinyin'] = st.text_input(
                f'拼音{w_i+1}', value=word.get('pinyin', ''),
                placeholder='如：lǎng rùn',
                key=f'py_{lec_i}_{w_i}', label_visibility='collapsed')
        with wc2:
            word['hanzi'] = st.text_input(
                f'汉字{w_i+1}', value=word.get('hanzi', ''),
                placeholder='如：朗润',
                key=f'hz_{lec_i}_{w_i}', label_visibility='collapsed')
        with wc3:
            if st.button('✕', key=f'del_word_{lec_i}_{w_i}'):
                words.pop(w_i)
                st.rerun()

    if not words:
        st.caption('（尚未添加词语）')

    if st.button('➕ 添加词语', key=f'add_word_{lec_i}'):
        words.append(_new_word())
        st.rerun()


# ══════════════════════════════════════════════════════════════
# 主界面
# ══════════════════════════════════════════════════════════════

def _check_password():
    if st.session_state.get('authenticated'):
        return True
    pwd = st.secrets.get('APP_PASSWORD', '')
    if not pwd:
        return True  # 未设置密码时直接放行
    st.title('📚 语文练习册生成器')
    entered = st.text_input('请输入访问密码', type='password', key='pwd_input')
    if st.button('进入'):
        if entered == pwd:
            st.session_state['authenticated'] = True
            st.rerun()
        else:
            st.error('密码错误，请重试')
    return False

if not _check_password():
    st.stop()

st.title('📚 语文练习册生成器')
st.caption('根据样章格式自动生成 Word 练习册，支持现代文阅读 / 作文 / 基础 / 文言文 / 名著 / 诗词')

if 'lectures' not in st.session_state:
    st.session_state.lectures = [_new_lecture(1)]

lectures = st.session_state.lectures

if st.button('➕ 新增一讲', use_container_width=False):
    lectures.append(_new_lecture(len(lectures) + 1))
    st.rerun()

st.divider()

# ── 遍历每讲 ──────────────────────────────────────────────────
for lec_i, lec in enumerate(lectures):
    subj_label  = lec.get('subject', SUBJECT_READING)
    topic_label = lec.get('topic') or '（未填写讲题）'
    with st.expander(
        f"**第 {lec['number']} 讲**  ｜  {subj_label}  ｜  {topic_label}",
        expanded=(lec_i == 0),
    ):
        if len(lectures) > 1:
            if st.button('🗑 删除本讲', key=f'del_lec_{lec_i}'):
                lectures.pop(lec_i)
                for i, l in enumerate(lectures):
                    l['number'] = i + 1
                st.rerun()

        # 基本信息
        col_num, col_subj, col_topic = st.columns([1, 2, 3])
        with col_num:
            lec['number'] = st.number_input(
                '讲次', min_value=1, value=lec['number'], key=f'num_{lec_i}')
        with col_subj:
            idx = ALL_SUBJECTS.index(lec['subject']) if lec['subject'] in ALL_SUBJECTS else 0
            lec['subject'] = st.selectbox(
                '板块类型', ALL_SUBJECTS, index=idx, key=f'subj_{lec_i}')
        with col_topic:
            lec['topic'] = st.text_input(
                '讲题', value=lec.get('topic', ''),
                placeholder='如：记叙文标题作用题', key=f'topic_{lec_i}')

        lec['core_method_lines'] = st.slider(
            '核心方法复习区行数', 4, 15,
            lec.get('core_method_lines', 8),
            key=f'cm_lines_{lec_i}',
            help='学生自主复习核心技法的空白书写区行数')

        st.divider()

        subject = lec['subject']

        if subject == SUBJECT_WRITING:
            st.info('📝 作文板块：核心方法 + 巩固提升（写作题）+ 能力进阶（写作题）+ 妙笔生花 + 文心雕龙')
            render_writing_tabs(lec_i, lec)
            st.markdown('> 将自动在文档末尾添加**妙笔生花**和**文心雕龙**自评表')

        elif subject == SUBJECT_BASIC:
            st.info('📝 基础板块：核心方法 + 巩固提升（题目）+ 能力进阶（题目）+ 日积月累（描红）')
            render_article_question_tabs(lec_i, lec)
            st.divider()
            render_daily_words(lec_i, lec)

        else:
            render_article_question_tabs(lec_i, lec)

# ── 生成 ──────────────────────────────────────────────────────
st.divider()
col_btn, _ = st.columns([2, 4])
with col_btn:
    generate = st.button('🎯 生成练习册 Word 文档',
                         use_container_width=True, type='primary')

if generate:
    data = {'lectures': st.session_state.lectures}
    with st.spinner('正在生成文档…'):
        try:
            buf = build_document(data)
            st.success('✅ 生成成功！')
            st.download_button(
                label='⬇️ 下载练习册 .docx',
                data=buf,
                file_name='练习册.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                use_container_width=True,
            )
        except Exception as e:
            st.error(f'生成失败：{e}')
            st.exception(e)
