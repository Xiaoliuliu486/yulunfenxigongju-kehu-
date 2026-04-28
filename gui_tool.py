import os
import sys
import json
import time
import queue
import threading
import webbrowser
from datetime import datetime
from collections import Counter

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

from excel_processor import read_excel
from scraper import scrape_url, get_analysis_text, detect_platform
from sentiment_analyzer import analyze_sentiment, classify_negative_reviews, generate_negative_summary


class AnalysisWorker(threading.Thread):
    def __init__(self, file_path, msg_queue):
        super().__init__(daemon=True)
        self.file_path = file_path
        self.msg_queue = msg_queue
        self.stop_flag = False

    def stop(self):
        self.stop_flag = True

    def run(self):
        try:
            self.msg_queue.put(('log', f'正在读取Excel文件...'))
            result = read_excel(self.file_path)

            if not result['success']:
                self.msg_queue.put(('error', f'读取失败: {result.get("error", "未知错误")}'))
                return

            reviews = result['reviews']
            if 'sheets' in result:
                sheets_text = []
                for s in result['sheets']:
                    sheets_text.append(f'  [{s["name"]}] {s["rows"]}行, {s["valid"]}条有效')
                    cols = s.get('columns', {})
                    sheets_text.append(f'    摘要列: {cols.get("摘要","?")}, 链接列: {cols.get("链接","?")}')
                self.msg_queue.put(('sheets_info', '\n'.join(sheets_text)))

            if not reviews:
                self.msg_queue.put(('error', '未找到有效数据！请确保Excel包含链接列'))
                return

            self.msg_queue.put(('log', f'共读取 {len(reviews)} 条，正在去重...'))

            seen_links = set()
            deduped = []
            dup_count = 0
            for r in reviews:
                link = r['link'].strip().rstrip('/')
                if link in seen_links:
                    dup_count += 1
                    continue
                seen_links.add(link)
                r['link'] = link
                deduped.append(r)
            reviews = deduped

            self.msg_queue.put(('dedup', dup_count))
            self.msg_queue.put(('log', f'去重完成，待分析: {len(reviews)} 条'))
            self.msg_queue.put(('total', len(reviews)))

            normal_count = sum(1 for r in reviews if '正常' in r.get('sheet', ''))
            if normal_count > 0:
                self.msg_queue.put(('log', f'发现 {normal_count} 条正常Sheet评价(跳过爬取)，{len(reviews)-normal_count} 条敏感Sheet评价(需爬取)'))
            else:
                self.msg_queue.put(('log', f'共 {len(reviews)} 条，正在分析...'))

            analyzed_reviews = []
            start_time = time.time()

            for i, review in enumerate(reviews, 1):
                if self.stop_flag:
                    self.msg_queue.put(('log', '用户中断分析'))
                    self.msg_queue.put(('interrupted', analyzed_reviews))
                    return

                is_normal_sheet = '正常' in review.get('sheet', '')
                elapsed = time.time() - start_time
                avg_time = elapsed / max(1, i - 1)
                remaining = avg_time * (len(reviews) - i + 1)

                status_msg = f'正在分析第 {i}/{len(reviews)} 条'
                if is_normal_sheet:
                    status_msg += ' [摘要模式]'
                else:
                    status_msg += f' [{detect_platform(review["link"])}]'

                self.msg_queue.put(('progress', i, len(reviews), remaining, status_msg))

                if is_normal_sheet:
                    scrape_result = {'scrape_success': False, 'platform': '', 'title': '', 'error': ''}
                    analysis_text = review.get('summary', '')
                else:
                    self.msg_queue.put(('url', review['link'][:70]))
                    scrape_result = scrape_url(review['link'], review.get('summary', ''))
                    analysis_text = get_analysis_text(scrape_result, review.get('summary', ''))

                sentiment = analyze_sentiment(analysis_text)
                platform = review.get('platform') or scrape_result.get('platform') or detect_platform(review['link'])

                analyzed = {
                    'row_index': review.get('row_index', i),
                    'summary': review.get('summary', ''),
                    'link': review['link'],
                    'platform': platform,
                    'sheet': review.get('sheet', ''),
                    'title': scrape_result.get('title', ''),
                    'scrape_success': scrape_result.get('scrape_success', False),
                    'scrape_error': scrape_result.get('error', ''),
                    'is_negative': sentiment['is_negative'],
                    'sentiment_score': sentiment['score'],
                    'negative_score': sentiment['negative_score'],
                    'negative_keywords': sentiment['negative_keywords'],
                    'positive_keywords': sentiment.get('positive_keywords', []),
                    'confidence': sentiment.get('confidence', 'neutral'),
                    'raw_negative_hits': sentiment.get('raw_negative_hits', 0),
                    'raw_positive_hits': sentiment.get('raw_positive_hits', 0),
                }
                analyzed_reviews.append(analyzed)

                confidence_tag = {
                    'high': '高可信', 'medium': '中可信', 'low': '低可信'
                }.get(sentiment.get('confidence', ''), '')
                if sentiment['is_negative']:
                    found_kws = [k for k in sentiment['negative_keywords'][:3] if not k.startswith('[模式]')]
                    kw_text = ','.join(found_kws) if found_kws else sentiment.get('negative_keywords', [''])[0]
                    self.msg_queue.put(('neg_found', confidence_tag, kw_text))
                else:
                    self.msg_queue.put(('normal',))

                if i % 50 == 0:
                    self.save_progress(analyzed_reviews)

            self.save_progress(analyzed_reviews)
            self.msg_queue.put(('done', analyzed_reviews, self.file_path))

        except Exception as e:
            self.msg_queue.put(('error', f'分析出错: {str(e)}'))

    def save_progress(self, analyzed):
        prog_file = self.file_path + '.progress.json'
        try:
            with open(prog_file, 'w', encoding='utf-8') as f:
                json.dump({'analyzed': len(analyzed), 'total': len(analyzed), 'reviews': analyzed}, f, ensure_ascii=False)
        except Exception:
            pass


class AnalysisGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('多平台舆论分析工具 v2.0')
        self.root.geometry('820x720')
        self.root.minsize(700, 650)

        try:
            self.root.iconbitmap(default='')
        except Exception:
            pass

        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value='准备就绪，请选择Excel文件')
        self.progress_var = tk.DoubleVar(value=0)
        self.worker = None
        self.msg_queue = queue.Queue()
        self.analyzed_reviews = []
        self.report_path = ''

        self._build_ui()
        self._poll_queue()

        if len(sys.argv) > 1:
            fp = sys.argv[1].strip().strip('"').strip("'")
            if os.path.exists(fp):
                self.file_path.set(fp)
                self._show_file_info()

    def _build_ui(self):
        self.root.configure(bg='#f5f5f5')

        header_frame = tk.Frame(self.root, bg='#2c3e50', height=70)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        tk.Label(header_frame, text='多平台舆论分析工具',
                 font=('Microsoft YaHei', 18, 'bold'), fg='white', bg='#2c3e50').pack(side='top', pady=(8, 0))
        tk.Label(header_frame, text='上传Excel → 自动分析 → 生成报告',
                 font=('Microsoft YaHei', 9), fg='#bdc3c7', bg='#2c3e50').pack(side='top')

        main = tk.Frame(self.root, bg='#f5f5f5', padx=20, pady=15)
        main.pack(fill='both', expand=True)

        file_frame = tk.LabelFrame(main, text=' 1. 选择Excel文件 ',
                                   font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#f5f5f5', padx=15, pady=12)
        file_frame.pack(fill='x', pady=(0, 10))

        path_row = tk.Frame(file_frame, bg='#f5f5f5')
        path_row.pack(fill='x')

        self.path_entry = tk.Entry(path_row, textvariable=self.file_path,
                                   font=('Microsoft YaHei', 10),
                                   relief='solid', bd=1, state='readonly')
        self.path_entry.pack(side='left', fill='x', expand=True, ipady=3)

        browse_btn = tk.Button(path_row, text='浏览...', font=('Microsoft YaHei', 10),
                               bg='#3498db', fg='white', relief='flat', padx=15,
                               activebackground='#2980b9', cursor='hand2',
                               command=self._browse_file)
        browse_btn.pack(side='right', padx=(10, 0))

        hint = tk.Label(file_frame, text='支持 .xlsx 格式，自动读取所有Sheet',
                        font=('Microsoft YaHei', 9), fg='#7f8c8d', bg='#f5f5f5')
        hint.pack(anchor='w', pady=(6, 0))

        info_frame = tk.LabelFrame(main, text=' 2. 文件信息 ',
                                   font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#f5f5f5', padx=15, pady=10)
        info_frame.pack(fill='x', pady=(0, 10))

        self.info_text = tk.Text(info_frame, height=4, font=('Microsoft YaHei', 9),
                                 relief='solid', bd=1, bg='#fafafa',
                                 state='disabled', wrap='word')
        self.info_text.pack(fill='x')
        self.info_text.tag_config('info', foreground='#2c3e50')
        self.info_text.tag_config('green', foreground='#27ae60')
        self.info_text.tag_config('yellow', foreground='#f39c12')

        ctrl_frame = tk.LabelFrame(main, text=' 3. 开始分析 ',
                                   font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#f5f5f5', padx=15, pady=12)
        ctrl_frame.pack(fill='x', pady=(0, 10))

        btn_row = tk.Frame(ctrl_frame, bg='#f5f5f5')
        btn_row.pack(fill='x')

        self.start_btn = tk.Button(btn_row, text='开始分析', font=('Microsoft YaHei', 13, 'bold'),
                                   bg='#27ae60', fg='white', relief='flat', padx=30, pady=5,
                                   activebackground='#219a52', cursor='hand2',
                                   command=self._start_analysis)
        self.start_btn.pack(side='left')

        self.stop_btn = tk.Button(btn_row, text='停止', font=('Microsoft YaHei', 11),
                                  bg='#e74c3c', fg='white', relief='flat', padx=20, pady=5,
                                  activebackground='#c0392b', cursor='hand2',
                                  state='disabled', command=self._stop_analysis)
        self.stop_btn.pack(side='left', padx=(10, 0))

        self.progress_bar = ttk.Progressbar(ctrl_frame, mode='determinate',
                                            variable=self.progress_var, length=500)
        self.progress_bar.pack(fill='x', pady=(10, 0))

        self.progress_label = tk.Label(ctrl_frame, textvariable=self.status_text,
                                       font=('Microsoft YaHei', 9), fg='#555',
                                       bg='#f5f5f5', anchor='w', wraplength=700)
        self.progress_label.pack(fill='x', pady=(4, 0))

        self.detail_label = tk.Label(ctrl_frame, text='', font=('Microsoft YaHei', 9),
                                     fg='#888', bg='#f5f5f5', anchor='w', wraplength=700)
        self.detail_label.pack(fill='x')

        result_frame = tk.LabelFrame(main, text=' 4. 分析结果 ',
                                     font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#f5f5f5', padx=15, pady=10)
        result_frame.pack(fill='both', expand=True)

        result_top = tk.Frame(result_frame, bg='#f5f5f5')
        result_top.pack(fill='x')

        self.result_text = tk.Text(result_top, height=6, font=('Microsoft YaHei', 10),
                                   relief='solid', bd=1, bg='#fafafa',
                                   state='disabled', wrap='word')
        self.result_text.pack(fill='x', expand=True)
        self.result_text.tag_config('bold', font=('Microsoft YaHei', 10, 'bold'))
        self.result_text.tag_config('red', foreground='#e74c3c')
        self.result_text.tag_config('green', foreground='#27ae60')
        self.result_text.tag_config('blue', foreground='#2980b9')

        btn_row2 = tk.Frame(result_frame, bg='#f5f5f5')
        btn_row2.pack(fill='x', pady=(8, 0))

        self.open_report_btn = tk.Button(btn_row2, text='打开报告', font=('Microsoft YaHei', 11),
                                         bg='#3498db', fg='white', relief='flat', padx=20, pady=4,
                                         activebackground='#2980b9', cursor='hand2',
                                         state='disabled', command=self._open_report)
        self.open_report_btn.pack(side='left')

        self.view_detail_btn = tk.Button(btn_row2, text='查看负面详情', font=('Microsoft YaHei', 11),
                                         bg='#e67e22', fg='white', relief='flat', padx=20, pady=4,
                                         activebackground='#d35400', cursor='hand2',
                                         state='disabled', command=self._show_negative_detail)
        self.view_detail_btn.pack(side='left', padx=(10, 0))

        self.open_folder_btn = tk.Button(btn_row2, text='打开报告文件夹', font=('Microsoft YaHei', 10),
                                         bg='#95a5a6', fg='white', relief='flat', padx=15, pady=4,
                                         activebackground='#7f8c8d', cursor='hand2',
                                         state='disabled', command=self._open_folder)
        self.open_folder_btn.pack(side='left', padx=(10, 0))

        footer = tk.Frame(self.root, bg='#ecf0f1', height=30)
        footer.pack(fill='x', side='bottom')
        tk.Label(footer, text='支持平台: 抖音 / 小红书 / 快手 / 今日头条 / 懂车帝 / 黑猫投诉 等',
                 font=('Microsoft YaHei', 8), fg='#95a5a6', bg='#ecf0f1').pack(side='right', padx=15)

    def _browse_file(self):
        fp = filedialog.askopenfilename(
            title='选择Excel文件',
            filetypes=[('Excel文件', '*.xlsx *.xls'), ('所有文件', '*.*')]
        )
        if fp:
            self.file_path.set(fp)
            self._show_file_info()

    def _show_file_info(self):
        fp = self.file_path.get().strip().strip('"').strip("'")
        if not fp or not os.path.exists(fp):
            return

        self._set_info_text('正在读取文件信息...\n')

        result = read_excel(fp)
        if not result['success']:
            self._append_info(f'读取失败: {result.get("error", "")}', 'yellow')
            return

        lines = [f'文件: {os.path.basename(fp)}']
        total_rows = result.get('total_rows', 0)
        valid_rows = result.get('valid_rows', len(result['reviews']))
        lines.append(f'总数据: {total_rows} 行 | 有效: {valid_rows} 条')

        if 'sheets' in result:
            lines.append(f'Sheet数量: {len(result["sheets"])}')
            for s in result['sheets']:
                lines.append(f'  [{s["name"]}] {s["rows"]}行 → {s["valid"]}条有效')

        if result.get('errors'):
            err_count = len(result['errors'])
            lines.append(f'  跳过 {err_count} 行 (缺少链接/无效数据)')

        self._set_info_text('')
        for line in lines:
            tag = 'info'
            if '有效' in line or '跳过' in line:
                tag = 'green'
            self._append_info(line + '\n', tag)

        self.start_btn.configure(state='normal')

    def _set_info_text(self, text):
        self.info_text.configure(state='normal')
        self.info_text.delete('1.0', 'end')
        self.info_text.insert('1.0', text)
        self.info_text.configure(state='disabled')

    def _append_info(self, text, tag='info'):
        self.info_text.configure(state='normal')
        self.info_text.insert('end', text, tag)
        self.info_text.see('end')
        self.info_text.configure(state='disabled')

    def _set_result_text(self, text):
        self.result_text.configure(state='normal')
        self.result_text.delete('1.0', 'end')
        self.result_text.insert('1.0', text)
        self.result_text.configure(state='disabled')

    def _append_result(self, text, tag=''):
        self.result_text.configure(state='normal')
        self.result_text.insert('end', text, tag)
        self.result_text.see('end')
        self.result_text.configure(state='disabled')

    def _start_analysis(self):
        fp = self.file_path.get().strip().strip('"').strip("'")
        if not fp or not os.path.exists(fp):
            messagebox.showwarning('提示', '请先选择Excel文件')
            return

        prog_file = fp + '.progress.json'
        if os.path.exists(prog_file):
            resume = messagebox.askyesno('发现上次进度',
                                         '检测到上次未完成的分析，是否继续？\n\n选"否"将重新开始分析')
            if resume:
                try:
                    with open(prog_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    self.analyzed_reviews = data.get('reviews', [])
                    self._set_result_text('')
                    self._append_result(f'已恢复上次进度: {len(self.analyzed_reviews)} 条\n', 'blue')
                except Exception:
                    pass

        self.start_btn.configure(state='disabled', bg='#95a5a6')
        self.stop_btn.configure(state='normal')
        self.open_report_btn.configure(state='disabled')
        self.view_detail_btn.configure(state='disabled')
        self.open_folder_btn.configure(state='disabled')
        self.progress_var.set(0)
        self.status_text.set('正在初始化...')
        self.detail_label.configure(text='')

        self.worker = AnalysisWorker(fp, self.msg_queue)
        self.worker.start()

    def _stop_analysis(self):
        if self.worker and self.worker.is_alive():
            self.worker.stop()
            self.status_text.set('正在停止...')
            self.stop_btn.configure(state='disabled')

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self._handle_msg(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_msg(self, msg):
        msg_type = msg[0]

        if msg_type == 'log':
            self.status_text.set(msg[1])

        elif msg_type == 'sheets_info':
            self._append_info('\n' + msg[1] + '\n', 'info')

        elif msg_type == 'dedup':
            self._append_info(f'去除重复: {msg[1]} 条\n', 'yellow')

        elif msg_type == 'total':
            self.progress_var.set(0)
            self._append_info(f'待分析: {msg[1]} 条\n', 'green')

        elif msg_type == 'progress':
            current, total, remaining, status = msg[1], msg[2], msg[3], msg[4]
            pct = (current / total) * 100 if total > 0 else 0
            self.progress_var.set(pct)
            self.status_text.set(status)
            if remaining > 0:
                mins = int(remaining // 60)
                secs = int(remaining % 60)
                if mins > 0:
                    self.detail_label.configure(text=f'预计剩余: {mins}分{secs}秒')
                else:
                    self.detail_label.configure(text=f'预计剩余: {secs}秒')
            else:
                self.detail_label.configure(text='')

        elif msg_type == 'url':
            self.detail_label.configure(text=f'正在抓取: {msg[1]}...')

        elif msg_type == 'neg_found':
            conf, kws = msg[1], msg[2]
            self.detail_label.configure(text=f'  发现负面 [{conf}]: {kws}')

        elif msg_type == 'normal':
            pass

        elif msg_type == 'done':
            self.analyzed_reviews = msg[1]
            file_path = msg[2]
            self._generate_report(self.analyzed_reviews, file_path)
            self.start_btn.configure(state='normal', bg='#27ae60')
            self.stop_btn.configure(state='disabled')

        elif msg_type == 'interrupted':
            self.analyzed_reviews = msg[1]
            self.start_btn.configure(state='normal', bg='#27ae60')
            self.stop_btn.configure(state='disabled')
            self.status_text.set(f'已中断，已分析 {len(self.analyzed_reviews)} 条')
            prog_file = self.file_path.get().strip().strip('"').strip("'") + '.progress.json'
            if os.path.exists(prog_file):
                self.detail_label.configure(text=f'进度已保存，下次可继续')

        elif msg_type == 'error':
            messagebox.showerror('错误', msg[1])
            self.status_text.set(f'错误: {msg[1]}')
            self.start_btn.configure(state='normal', bg='#27ae60')
            self.stop_btn.configure(state='disabled')
            self.detail_label.configure(text='')

    def _generate_report(self, analyzed_reviews, file_path):
        total = len(analyzed_reviews)
        negative_reviews = classify_negative_reviews(analyzed_reviews)
        negative_count = len(negative_reviews)
        negative_pct = negative_count / total * 100 if total > 0 else 0

        platform_stats = Counter(r['platform'] for r in analyzed_reviews)
        negative_platform_stats = Counter(r['platform'] for r in negative_reviews)
        confidence_levels = Counter(r.get('confidence', 'neutral') for r in negative_reviews)
        positive_keywords_all = []
        for r in analyzed_reviews:
            positive_keywords_all.extend(r.get('positive_keywords', []))

        output_dir = os.path.dirname(file_path)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        self.report_path = os.path.join(output_dir, f'{base_name}_分析报告_{timestamp}.txt')

        summary_text = generate_negative_summary(negative_reviews)

        report_lines = []
        report_lines.append(f'{"="*60}')
        report_lines.append(f'  多平台舆论分析报告 v2.0')
        report_lines.append(f'  文件: {os.path.basename(file_path)}')
        report_lines.append(f'  分析时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        report_lines.append(f'{"="*60}')
        report_lines.append('')
        report_lines.append(f'[分析概览]')
        report_lines.append(f'  - 总评价数: {total}')
        report_lines.append(f'  - 负面评价: {negative_count} ({negative_pct:.1f}%)')
        report_lines.append(f'     - 高可信: {confidence_levels.get("high", 0)} 条')
        report_lines.append(f'     - 中可信: {confidence_levels.get("medium", 0)} 条')
        report_lines.append(f'     - 低可信: {confidence_levels.get("low", 0)} 条')
        report_lines.append(f'  - 正常评价: {total - negative_count}')
        report_lines.append('')
        report_lines.append(f'[各平台分布]')
        for p, c in platform_stats.most_common():
            neg_c = negative_platform_stats.get(p, 0)
            report_lines.append(f'  {p:　<8} {c:>4}条  负面{neg_c}条')
        report_lines.append('')
        report_lines.append(f'[高频正面词]')
        if positive_keywords_all:
            top_pos = Counter(positive_keywords_all).most_common(10)
            report_lines.append('  ' + ', '.join(f'"{kw}"({c}次)' for kw, c in top_pos))
        report_lines.append('')
        report_lines.append(f'{"-"*60}')
        report_lines.append('')
        report_lines.append(summary_text)
        report_lines.append('')
        report_lines.append(f'{"-"*60}')
        report_lines.append('')
        report_lines.append(f'[全部评价详细结果]')
        report_lines.append('')
        for i, r in enumerate(analyzed_reviews, 1):
            tag = '[负面]' if r['is_negative'] else '[正常]'
            conf = {'high': '[高可信]', 'medium': '[中可信]', 'low': '[低可信]'}.get(r.get('confidence', ''), '')
            report_lines.append(f'  {i}. {tag}{conf} {r["platform"]}')
            report_lines.append(f'     摘要: {r["summary"][:120] if r["summary"] else "无摘要"}')
            report_lines.append(f'     链接: {r["link"]}')
            if r['is_negative']:
                report_lines.append(f'     负面指数: {r["negative_score"]:.2f} | 可信度: {r.get("confidence", "?")}')
                if r['negative_keywords']:
                    kws = [k for k in r['negative_keywords'] if not k.startswith('[模式]')][:5]
                    if kws:
                        report_lines.append(f'     匹配关键词: {", ".join(kws)}')
            report_lines.append('')

        with open(self.report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))

        prog_file = file_path + '.progress.json'
        if os.path.exists(prog_file):
            os.remove(prog_file)

        self.progress_var.set(100)
        self.status_text.set(f'分析完成！共 {total} 条，负面 {negative_count} 条 ({negative_pct:.1f}%)')

        self._set_result_text('')
        self._append_result(f'分析完成！\n\n', 'bold')
        self._append_result(f'总评价: {total} 条\n')
        self._append_result(f'负面评价: {negative_count} 条 ({negative_pct:.1f}%)\n', 'red')
        self._append_result(f'  高可信: {confidence_levels.get("high", 0)}  |  ', 'red')
        self._append_result(f'中可信: {confidence_levels.get("medium", 0)}  |  ')
        self._append_result(f'低可信: {confidence_levels.get("low", 0)}\n')
        self._append_result(f'正常评价: {total - negative_count} 条\n\n', 'green')
        self._append_result(f'报告已保存至:\n')
        self._append_result(f'{self.report_path}\n', 'blue')

        self.open_report_btn.configure(state='normal')
        self.view_detail_btn.configure(state='normal' if negative_count > 0 else 'disabled')
        self.open_folder_btn.configure(state='normal')
        self.detail_label.configure(text='')

    def _open_report(self):
        if self.report_path and os.path.exists(self.report_path):
            os.startfile(self.report_path)

    def _open_folder(self):
        if self.report_path:
            folder = os.path.dirname(self.report_path)
            os.startfile(folder)

    def _show_negative_detail(self):
        negative_reviews = classify_negative_reviews(self.analyzed_reviews)
        if not negative_reviews:
            messagebox.showinfo('提示', '没有负面评价')
            return

        win = tk.Toplevel(self.root)
        win.title(f'负面评价详情 ({len(negative_reviews)} 条)')
        win.geometry('800x600')
        win.minsize(600, 400)

        header = tk.Frame(win, bg='#e74c3c', height=40)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text=f'负面评价列表 - 共 {len(negative_reviews)} 条',
                 font=('Microsoft YaHei', 12, 'bold'), fg='white', bg='#e74c3c').pack(pady=6)

        canvas = tk.Canvas(win, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='white')

        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw', width=780)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        scrollbar.pack(side='right', fill='y', pady=10)

        conf_colors = {'high': '#e74c3c', 'medium': '#e67e22', 'low': '#f1c40f'}
        conf_text = {'high': '高可信', 'medium': '中可信', 'low': '低可信'}

        for i, r in enumerate(negative_reviews, 1):
            conf = r.get('confidence', 'medium')
            card = tk.Frame(scroll_frame, bg='#fef9f9', relief='solid', bd=1, padx=12, pady=8)
            card.pack(fill='x', pady=4)

            row1 = tk.Frame(card, bg='#fef9f9')
            row1.pack(fill='x')
            tk.Label(row1, text=f'#{i}', font=('Microsoft YaHei', 9, 'bold'),
                     fg=conf_colors.get(conf, '#555'), bg='#fef9f9').pack(side='left')
            tk.Label(row1, text=f'[{conf_text.get(conf, conf)}]', font=('Microsoft YaHei', 9, 'bold'),
                     fg=conf_colors.get(conf, '#555'), bg='#fef9f9').pack(side='left', padx=(5, 0))
            tk.Label(row1, text=f'[{r["platform"]}]', font=('Microsoft YaHei', 9),
                     fg='#2980b9', bg='#fef9f9').pack(side='left', padx=(5, 0))
            if r.get('negative_score', 0) > 0:
                tk.Label(row1, text=f'负面指数: {r["negative_score"]:.2f}',
                         font=('Microsoft YaHei', 9), fg='#888', bg='#fef9f9').pack(side='right')

            summary = r.get('summary', '')[:200]
            if summary:
                tk.Label(card, text=summary, font=('Microsoft YaHei', 9),
                         fg='#333', bg='#fef9f9', wraplength=700, anchor='w',
                         justify='left').pack(fill='x', pady=(4, 2))

            link_frame = tk.Frame(card, bg='#fef9f9')
            link_frame.pack(fill='x')

            link_label = tk.Label(link_frame, text=r['link'][:80], font=('Microsoft YaHei', 8),
                                  fg='#3498db', bg='#fef9f9', cursor='hand2')
            link_label.pack(side='left')
            link_label.bind('<Button-1>', lambda e, url=r['link']: webbrowser.open(url))

            if r.get('negative_keywords'):
                kws = [k for k in r['negative_keywords'][:6] if not k.startswith('[模式]')]
                if kws:
                    tk.Label(card, text='关键词: ' + ', '.join(kws), font=('Microsoft YaHei', 8),
                             fg='#e67e22', bg='#fef9f9').pack(anchor='w', pady=(2, 0))

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = AnalysisGUI()
    app.run()
