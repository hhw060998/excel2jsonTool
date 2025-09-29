#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的图形界面，用来配置导表所需的路径和 Python 可执行程序，
并在配置好后生成或运行批处理文件以启动现有导表流程。

用法：在 Windows 上用系统 Python 运行：
    python tools/gui_export_launcher.py

功能：
- 选择 Python 可执行文件、Excel 根目录、项目 JSON 输出、C# 脚本输出、枚举输出。
- 验证路径的存在性与可执行性（python）。
- 将配置保存到仓库根目录下的 `gui_export_config.json`。
- 生成一个批处理文件 `导表_from_gui.bat` 放到 Excel 根目录（备份原有文件）。
- 运行生成的批处理并把输出实时显示到窗口中。

设计原则：不改动原导表工具代码，仅生成/运行批处理来调用它。
"""

import json
import os
import re
import threading
import subprocess
import locale
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG = ROOT / 'gui_export_config.json'

DEFAULT_BATCH_NAME = '导表_from_gui.bat'


class LauncherApp:
    def __init__(self, master):
        self.master = master
        master.title('导表')

        self.cfg = {
            'python_exe': '',
            'excel_root': '',
            'output_project': '',
            'cs_output': '',
            'enum_output': '',
            'batch_path': '',
        }

        # 状态标签集合，用于显示每个路径是否有效
        self.status_labels = {}

        self._build_ui()
        # 记录加载时配置，用于判断是否有未保存更改
        self._loaded_cfg = dict(self.cfg)
        self._dirty = False
        self.proc = None

        self._load_config()

    def _build_ui(self):
        row = 0
        self._original_paths = {}
        self._reset_buttons = {}
        def add_entry(label_text, key, browse_mode='file'):
            nonlocal row
            lbl = tk.Label(self.master, text=label_text)
            lbl.grid(row=row, column=0, sticky='w', padx=6, pady=4)
            ent = tk.Entry(self.master, width=60, state='readonly')
            ent.grid(row=row, column=1, padx=6, pady=4)
            btn = tk.Button(self.master, text='...', width=3,
                            command=(lambda k=key, e=ent, m=browse_mode: self._on_browse(k, e, m)))
            btn.grid(row=row, column=2, padx=(2,2), pady=4)
            reset_btn = tk.Button(self.master, text='重置', width=4,
                                  command=(lambda k=key, e=ent: self._on_reset_path(k, e)))
            reset_btn.grid(row=row, column=3, padx=(2,2), pady=4)
            reset_btn.grid_remove()
            self._reset_buttons[key] = reset_btn
            st = tk.Label(self.master, text='', fg='red')
            st.grid(row=row, column=4, padx=6, pady=4, sticky='w')
            self.status_labels[key] = st
            row += 1
            return ent

        self.e_python = add_entry('Python 可执行文件:', 'python_exe', 'file')
        self.e_excel = add_entry('Excel 根目录:', 'excel_root', 'dir')
        self.e_output_project = add_entry('工程 JSON 输出目录:', 'output_project', 'dir')
        self.e_cs_output = add_entry('C# 脚本输出目录:', 'cs_output', 'dir')
        self.e_enum_output = add_entry('枚举输出目录:', 'enum_output', 'dir')

        lbl = tk.Label(self.master, text='批处理文件（可选）:')
        lbl.grid(row=row, column=0, sticky='w', padx=6, pady=4)
        self.e_batch = tk.Entry(self.master, width=60, state='readonly')
        self.e_batch.grid(row=row, column=1, padx=6, pady=4)
        btnb = tk.Button(self.master, text='...', width=3, command=self._browse_batch)
        btnb.grid(row=row, column=2, padx=(2,2), pady=4)
        st = tk.Label(self.master, text='', fg='red')
        st.grid(row=row, column=4, padx=6, pady=4, sticky='w')
        self.status_labels['batch_path'] = st
        row += 1

        frm = tk.Frame(self.master)
        frm.grid(row=row, column=0, columnspan=3, pady=6)
        self.btn_validate = tk.Button(frm, text='校验', command=self._validate)
        self.btn_validate.grid(row=0, column=0, padx=4)
        self.btn_save = tk.Button(frm, text='保存目录', command=self._on_save_and_refresh)
        self.btn_save.grid(row=0, column=1, padx=4)
        self.btn_run = tk.Button(frm, text='导表', command=self._run_batch)
        self.btn_run.grid(row=0, column=2, padx=4)
        self._button_frame = frm
        row += 1

        lbl_out = tk.Label(self.master, text='运行输出:')
        lbl_out.grid(row=row, column=0, sticky='w', padx=6, pady=4)
        row += 1
        # 让日志文本区随窗口拉伸
        self.txt = scrolledtext.ScrolledText(self.master, height=20, width=92)
        self.txt.grid(row=row, column=0, columnspan=5, padx=6, pady=4, sticky='nsew')

        # 配置grid行列权重，只有日志区所在行/列可扩展
        for i in range(row):
            self.master.grid_rowconfigure(i, weight=0)
        self.master.grid_rowconfigure(row, weight=1)
        for i in range(5):
            self.master.grid_columnconfigure(i, weight=1 if i == 1 else 0)

    def _on_browse(self, key, entry_widget, mode):
        # 如果当前字段已有有效路径，使用它作为初始目录；否则使用仓库根或用户目录
        cur = entry_widget.get().strip()
        if cur and os.path.exists(cur):
            init_dir = cur if os.path.isdir(cur) else str(Path(cur).parent)
        else:
            init_dir = str(ROOT)

        if mode == 'file':
            p = filedialog.askopenfilename(initialdir=init_dir)
        else:
            p = filedialog.askdirectory(initialdir=init_dir)
        if p:
            self._set_entry_value(entry_widget, p)
            # 标记已修改并重新校验
            self.cfg[key] = p
            self._validate_and_show_status(key)
            # 标记为脏
            self._dirty = any(self.cfg.get(k) != self._loaded_cfg.get(k) for k in ('python_exe', 'excel_root', 'output_project', 'cs_output', 'enum_output', 'batch_path'))
            if key in self._reset_buttons:
                if self.cfg.get(key) != self._original_paths.get(key, ''):
                    self._reset_buttons[key].grid()
                else:
                    self._reset_buttons[key].grid_remove()
            self._refresh_action_buttons()

    def _set_entry_value(self, entry_widget, value: str):
        # 由于 Entry 设为 readonly，临时切换状态以写入值
        try:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, value)
        finally:
            try:
                entry_widget.config(state='readonly')
            except Exception:
                pass

    def _browse_batch(self):
        # 使用当前 batch 字段作为起始目录
        cur = self.e_batch.get().strip()
        if cur and os.path.exists(cur):
            init_dir = cur if os.path.isdir(cur) else str(Path(cur).parent)
        else:
            init_dir = str(ROOT)

        p = filedialog.askopenfilename(initialdir=init_dir, filetypes=[('Batch File', '*.bat'), ('All files', '*.*')])
        if p:
            self._set_entry_value(self.e_batch, p)
            self.cfg['batch_path'] = p
            self._validate_and_show_status('batch_path')
            self._dirty = any(self.cfg.get(k) != self._loaded_cfg.get(k) for k in ('python_exe', 'excel_root', 'output_project', 'cs_output', 'enum_output', 'batch_path'))

    def _load_config(self):
        if DEFAULT_CONFIG.exists():
            try:
                d = json.loads(DEFAULT_CONFIG.read_text(encoding='utf-8'))
                self.cfg.update(d)
            except Exception:
                pass
        # 若 JSON 配置不存在或部分字段未设置，尝试从常见批处理文件解析配置（优先）
        if not self.cfg.get('python_exe') or not self.cfg.get('excel_root'):
            # 在多个目录查找 .bat 文件：优先使用配置中指定的 batch 路径所在目录，再尝试当前工作目录、脚本所在目录、仓库 ExcelFolder
            search_dirs = []
            # 如果配置里有 batch_path，则先使用其目录
            if self.cfg.get('batch_path'):
                try:
                    p = Path(self.cfg['batch_path'])
                    if p.parent.exists():
                        search_dirs.append(p.parent)
                except Exception:
                    pass
            # 当前工作目录（即双击运行批处理所在目录）
            search_dirs.append(Path.cwd())
            # 脚本所在目录（tools）
            script_dir = Path(__file__).resolve().parent
            search_dirs.append(script_dir)
            # 仓库的 ExcelFolder（原始位置）
            search_dirs.append(ROOT / 'ExcelFolder')

            # 收集候选 .bat（按目录顺序去重）
            seen = set()
            candidates = []
            for d in search_dirs:
                try:
                    if not d or not d.exists():
                        continue
                    # 优先尝试特定文件名
                    preferred = d / '!【导表】.bat'
                    if preferred.exists() and str(preferred) not in seen:
                        candidates.append(preferred)
                        seen.add(str(preferred))
                    for b in d.glob('*.bat'):
                        sb = str(b)
                        if sb not in seen:
                            candidates.append(b)
                            seen.add(sb)
                except Exception:
                    continue

            for bat in candidates:
                try:
                    parsed = self._parse_batch_file(bat)
                    if parsed:
                        # 填充未设置或覆盖空值
                        for k, v in parsed.items():
                            if v and (not self.cfg.get(k)):
                                self.cfg[k] = v
                        # 记录来源批处理文件
                        self.cfg.setdefault('_parsed_from', str(bat))
                        # 如果已尽可能填充则停止
                        if self.cfg.get('excel_root') and self.cfg.get('output_project'):
                            break
                except Exception:
                    # 忽略单个批处理解析错误
                    pass

        # 填充 UI
        for k, e in (('python_exe', self.e_python), ('excel_root', self.e_excel),
                     ('output_project', self.e_output_project), ('cs_output', self.e_cs_output),
                     ('enum_output', self.e_enum_output), ('batch_path', self.e_batch)):
            v = self.cfg.get(k) or ''
            self._set_entry_value(e, v)
            self._original_paths[k] = v
            # 初始时隐藏重置按钮
            if hasattr(self, '_reset_buttons') and k in self._reset_buttons:
                self._reset_buttons[k].grid_remove()

        # 绑定关闭事件以确保可清理子进程
        try:
            self.master.protocol('WM_DELETE_WINDOW', self._on_close)
        except Exception:
            pass
        # 打开时立即校验并显示状态
        self._validate_and_show_status()
        # 新增：首次加载后刷新按钮区，避免初始时两个按钮都显示
        if hasattr(self, '_refresh_action_buttons'):
            self._refresh_action_buttons()

    def _on_field_changed(self, key: str):
        # 更新 cfg，并比较是否与加载时相同
        try:
            self.cfg[key] = getattr(self, f'e_{key}').get().strip()
        except Exception:
            # fallback: read directly from entries dict mapping
            try:
                # some keys don't have e_<key> name (output_project uses e_output_project)
                val = self.e_batch.get().strip() if key == 'batch_path' else None
                if val is not None:
                    self.cfg[key] = val
            except Exception:
                pass
        # 比较当前 cfg 与加载时的 cfg
        self._dirty = any(self.cfg.get(k) != self._loaded_cfg.get(k) for k in ('python_exe', 'excel_root', 'output_project', 'cs_output', 'enum_output', 'batch_path'))
        # 重新校验当前字段状态显示
        self._validate_and_show_status(key)

    def _validate_and_show_status(self, specific_key: str | None = None):
        import shutil
        def check_python(val: str) -> bool:
            if not val:
                return False
            if val.lower() == 'python':
                return shutil.which('python') is not None
            return Path(val).exists()

        def check_dir(val: str) -> bool:
            if not val:
                return False
            return Path(val).is_dir()

        checks = {
            'python_exe': check_python,
            'excel_root': check_dir,
            'output_project': check_dir,
            'cs_output': check_dir,
            'enum_output': check_dir,
            'batch_path': lambda v: Path(v).exists(),
        }

        keys = [specific_key] if specific_key else list(self.status_labels.keys())
        for k in keys:
            lbl = self.status_labels.get(k)
            if not lbl:
                continue
            val = self.cfg.get(k) or ''
            ok = checks.get(k, lambda v: False)(val)
            # 新增：导出路径不含Assets时高亮黄色tip
            if k in ('output_project', 'cs_output', 'enum_output') and ok:
                if 'assets' not in str(val).replace('\\', '/').lower():
                    lbl.config(text='非Unity目录', fg='orange')
                else:
                    lbl.config(text='有效', fg='green')
            else:
                if ok:
                    lbl.config(text='有效', fg='green')
                else:
                    if not val:
                        lbl.config(text='空', fg='orange')
                    else:
                        lbl.config(text='无效', fg='red')

    def _parse_batch_file(self, path: Path) -> dict:
        """解析批处理文件中的 set 变量或 direct python 调用，返回映射到 GUI 字段的字典。

        支持解析形式:
         set input_folder=...\n set output_project_folder=...\n 等
        也会尝试解析调用行中带的 python 可执行路径（包含双引号）和 export_all.py 路径。
        返回键: python_exe, excel_root, output_project, cs_output, enum_output, batch_path
        """
        enc = locale.getpreferredencoding(False) or 'utf-8'
        text = path.read_text(encoding=enc, errors='ignore')
        res: dict = {}
        # 匹配 set VAR=VALUE（忽略大小写和空格）
        for m in re.finditer(r'(?im)^\s*set\s+([A-Za-z0-9_]+)\s*=\s*(.*)$', text):
            name = m.group(1).strip()
            val = m.group(2).strip().strip('"')
            # 映射常见变量名到 cfg key
            if name.lower() == 'input_folder':
                res['excel_root'] = val
            elif name.lower() == 'output_project_folder' or name.lower() == 'output_project':
                res['output_project'] = val
            elif name.lower() == 'csfile_output_folder' or name.lower() == 'cs_output_folder':
                res['cs_output'] = val
            elif name.lower() == 'enum_output_folder' or name.lower() == 'enum_output':
                res['enum_output'] = val

        # 尝试从调用 export_all.py 的行中提取 python 可执行路径
        for line in text.splitlines():
            if 'export_all.py' in line:
                # 找到第一对双引号里的可执行路径或第一个 token
                qm = re.search(r'"([^"]*python(?:\.exe)?)"', line, flags=re.I)
                if qm:
                    res['python_exe'] = qm.group(1)
                else:
                    # 如果没有双引号，尝试取第一个 token
                    tokens = line.strip().split()
                    if tokens:
                        first = tokens[0].strip('"')
                        if 'python' in first.lower():
                            res['python_exe'] = first
                break

        # 记录 batch 路径
        res['batch_path'] = str(path)
        return res

    def _save_config(self):
        self._read_ui_to_cfg()
        # 确定要写入的批处理文件：优先使用用户指定的 batch_path，其次使用解析来源
        batch_path = None
        if self.cfg.get('batch_path'):
            batch_path = Path(self.cfg['batch_path'])
        elif self.cfg.get('_parsed_from'):
            batch_path = Path(self.cfg['_parsed_from'])

        if not batch_path or not batch_path.exists():
            # 让用户选择要写入的批处理文件
            res = filedialog.asksaveasfilename(title='选择要写入的批处理文件或新建', defaultextension='.bat', filetypes=[('Batch File', '*.bat')])
            if not res:
                return
            batch_path = Path(res)

        # 确认提示
        ok = messagebox.askyesno('确认写入', f'将把当前配置写入批处理文件: {batch_path}\n(原文件会备份为 .bak)。继续吗？')
        if not ok:
            return

        # 备份原文件
        try:
            if batch_path.exists():
                bak = batch_path.with_suffix(batch_path.suffix + '.bak')
                try:
                    batch_path.replace(bak)
                except Exception:
                    try:
                        batch_path.rename(bak)
                    except Exception:
                        pass
        except Exception:
            # 忽略备份错误
            pass

        # 写入批处理（替换原有内容为标准模板）
        enc = locale.getpreferredencoding(False) or 'utf-8'
        export_all = (Path(__file__).resolve().parents[1] / 'ExcelExportTool' / 'export_all.py').resolve()
        content_lines = [
            '@echo off',
            'REM 由 gui_export_launcher 自动生成/更新，请确认路径正确',
            f'set input_folder={self.cfg.get("excel_root", "")}',
            f'set output_project_folder={self.cfg.get("output_project", "")}',
            f'set csfile_output_folder={self.cfg.get("cs_output", "")}',
            f'set enum_output_folder={self.cfg.get("enum_output", "")}',
            '',
            'REM 运行 python 脚本',
            f'"{self.cfg.get("python_exe", "python")}" "{export_all}" "%input_folder%" "%output_project_folder%" "%csfile_output_folder%" "%enum_output_folder%"',
            'pause',
        ]
        try:
            batch_path.write_text('\n'.join(content_lines), encoding=enc)
            messagebox.showinfo('保存成功', f'已写入: {batch_path}')
            # 更新 UI 中的 batch_path
            self._set_entry_value(self.e_batch, str(batch_path))
            self.cfg['batch_path'] = str(batch_path)
            self.cfg['_parsed_from'] = str(batch_path)
        except Exception as e:
            messagebox.showerror('写入失败', str(e))

    def _read_ui_to_cfg(self):
        self.cfg['python_exe'] = self.e_python.get().strip()
        self.cfg['excel_root'] = self.e_excel.get().strip()
        self.cfg['output_project'] = self.e_output_project.get().strip()
        self.cfg['cs_output'] = self.e_cs_output.get().strip()
        self.cfg['enum_output'] = self.e_enum_output.get().strip()
        self.cfg['batch_path'] = self.e_batch.get().strip()

    def _validate(self):
        self._read_ui_to_cfg()
        errors = []
        if not self.cfg['python_exe']:
            errors.append('未指定 Python 可执行文件')
        else:
            p = Path(self.cfg['python_exe'])
            if not p.exists():
                errors.append('Python 可执行文件不存在')
        if not self.cfg['excel_root']:
            errors.append('未指定 Excel 根目录')
        else:
            if not Path(self.cfg['excel_root']).is_dir():
                errors.append('Excel 根目录不存在')
        # 输出目录可不存在（生成时会创建），但路径应为非空字符串
        for k in ('output_project', 'cs_output', 'enum_output'):
            if not self.cfg[k]:
                errors.append(f'未指定 {k}')

        if errors:
            messagebox.showerror('校验失败', '\n'.join(errors))
            return False
        else:
            messagebox.showinfo('校验', '校验通过')
            return True

    def _generate_batch(self):
        if not self._validate():
            return
        self._read_ui_to_cfg()
        excel_root = Path(self.cfg['excel_root'])
        batch_path = Path(excel_root) / DEFAULT_BATCH_NAME
        # 如果用户指定了批处理路径则使用用户指定路径
        if self.cfg.get('batch_path'):
            batch_path = Path(self.cfg['batch_path'])
        # 备份已有批处理
        if batch_path.exists():
            bak = batch_path.with_suffix(batch_path.suffix + '.bak')
            try:
                batch_path.replace(bak)
            except Exception:
                try:
                    batch_path.rename(bak)
                except Exception:
                    pass
        # 找到 export_all.py 的绝对路径（假设仓库结构不变）
        export_all = (Path(__file__).resolve().parents[1] / 'ExcelExportTool' / 'export_all.py').resolve()
        content_lines = [
            '@echo off',
            'REM 由 gui_export_launcher 自动生成，请确认路径正确',
            f'set input_folder={self.cfg["excel_root"]}',
            f'set output_project_folder={self.cfg["output_project"]}',
            f'set csfile_output_folder={self.cfg["cs_output"]}',
            f'set enum_output_folder={self.cfg["enum_output"]}',
            '',
            'REM 运行 python 脚本',
            f'"{self.cfg["python_exe"]}" "{export_all}" "%input_folder%" "%output_project_folder%" "%csfile_output_folder%" "%enum_output_folder%"',
            'pause',
        ]
        try:
            batch_path.write_text('\n'.join(content_lines), encoding='utf-8')
            messagebox.showinfo('生成批处理', f'已生成: {batch_path}')
            self._set_entry_value(self.e_batch, str(batch_path))
            self.cfg['batch_path'] = str(batch_path)
        except Exception as e:
            messagebox.showerror('生成失败', str(e))

    def _run_batch(self):
        # 确保存在批处理
        self._read_ui_to_cfg()
        batch = self.cfg.get('batch_path')
        if not batch:
            messagebox.showerror('运行失败', '未指定要运行的批处理文件，请先生成或选择批处理文件')
            return
        batch_path = Path(batch)
        if not batch_path.exists():
            messagebox.showerror('运行失败', f'批处理不存在: {batch_path}')
            return

        # 检查所有导出路径是否包含Assets
        def _contains_assets(*paths):
            for p in paths:
                if p and 'assets' in str(p).replace('\\', '/').lower():
                    return True
            return False

        need_confirm = not _contains_assets(self.cfg.get('output_project'), self.cfg.get('cs_output'), self.cfg.get('enum_output'))
        if need_confirm:
            ok = messagebox.askyesno('非Unity目录', '导出路径不包含“Assets”，这通常不是Unity项目目录，是否继续导出？')
            if not ok:
                return
        # 一律在 GUI 中捕获输出并显示
        self.btn_run.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)
        self.txt.delete('1.0', tk.END)

        def target():
            try:
                # 若已确认，追加 --force-no-assets 参数
                extra_args = []
                if need_confirm:
                    extra_args.append('--force-no-assets')
                cmd = ['cmd.exe', '/c', str(batch_path)] + extra_args
                self._append_text(f'运行: {" ".join(cmd)}\n')
                enc = locale.getpreferredencoding(False) or 'utf-8'
                import os
                env = os.environ.copy()
                env['SHEETEASE_GUI'] = '1'
                proc = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True, encoding=enc, env=env)
                self.proc = proc
                for line in proc.stdout:
                    self._append_text(line)
                proc.wait()
                self._append_text(f"\n进程退出，返回码: {proc.returncode}\n")
            except Exception as e:
                self._append_text(f"运行失败: {e}\n")
            finally:
                self.proc = None
                try:
                    self.btn_run.config(state=tk.NORMAL)
                    self.btn_save.config(state=tk.NORMAL)
                except Exception:
                    pass

        t = threading.Thread(target=target, daemon=True)
        t.start()

    def _on_close(self):
        # 关闭窗口前尝试终止正在运行的子进程（无论是否为分离模式）
        try:
            if getattr(self, 'proc', None) is not None:
                try:
                    self.proc.kill()
                except Exception:
                    try:
                        # fallback: terminate
                        self.proc.terminate()
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self.master.destroy()
        except Exception:
            pass

    def _append_text(self, s: str):
        # 去除ANSI颜色码，防止GUI日志乱码
        import re
        ansi_escape = re.compile(r'\x1b\[[0-9;]*m')
        clean_s = ansi_escape.sub('', s)
        self.txt.insert(tk.END, clean_s)
        self.txt.see(tk.END)

    def _on_reset_path(self, key, entry_widget):
        orig = self._original_paths.get(key, '')
        self._set_entry_value(entry_widget, orig)
        self.cfg[key] = orig
        self._validate_and_show_status(key)
        if key in self._reset_buttons:
            self._reset_buttons[key].grid_remove()
        self._dirty = any(self.cfg.get(k) != self._loaded_cfg.get(k) for k in ('python_exe', 'excel_root', 'output_project', 'cs_output', 'enum_output', 'batch_path'))
        self._refresh_action_buttons()

    def _refresh_action_buttons(self):
        # 有未保存修改时只显示保存目录，否则只显示运行批处理
        if self._dirty:
            self.btn_save.grid()
            self.btn_run.grid_remove()
        else:
            self.btn_save.grid_remove()
            self.btn_run.grid()

    def _on_save_and_refresh(self):
        self._save_config()
        # 保存后刷新原始路径和按钮区
        for k in self._original_paths:
            self._original_paths[k] = self.cfg.get(k, '')
            if k in self._reset_buttons:
                self._reset_buttons[k].grid_remove()
        self._dirty = False
        self._refresh_action_buttons()

def main():
    root = tk.Tk()
    app = LauncherApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
