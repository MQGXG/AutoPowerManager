using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Security.Principal;
using Microsoft.Win32.TaskScheduler;

namespace AutoPowerManager
{
    public partial class MainForm : Form
    {
        // Tab控件
        private TabControl tabControl;
        private TabPage tabScheduled;
        private TabPage tabImmediate;
        private TabPage tabStageTime;

        // 定时任务控件
        private GroupBox grpBoot;
        private Label lblBootTime;
        private DateTimePicker dtpBootDate;
        private DateTimePicker dtpBootTime;
        private CheckBox chkBootEnabled;
        private CheckBox chkBootRepeat;
        private Button btnSetBoot;
        private Button btnRemoveBoot;

        private GroupBox grpShutdown;
        private Label lblShutdownTime;
        private DateTimePicker dtpShutdownDate;
        private DateTimePicker dtpShutdownTime;
        private CheckBox chkShutdownEnabled;
        private Button btnSetShutdown;
        private Button btnCancelShutdown;

        // 立即操作控件
        private GroupBox grpImmediate;
        private Button btnShutdownNow;
        private Button btnRestartNow;
        private Button btnCancelNow;
        private Label lblStatus;

        // 阶段时间控件
        private GroupBox grpStageSettings;
        private CheckedListBox chkListDays;
        private Label lblStageBootTime;
        private DateTimePicker dtpStageBootTime;
        private Label lblStageShutdownTime;
        private DateTimePicker dtpStageShutdownTime;
        private CheckBox chkStageEnabled;
        private Button btnSetStageTask;
        private Button btnRemoveStageTask;
        private ListView listStageTasks;

        // 系统托盘
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem menuShow;
        private ToolStripMenuItem menuExit;

        // 颜色方案
        private Color primaryColor = Color.FromArgb(41, 128, 185);   // 蓝色主题
        private Color secondaryColor = Color.FromArgb(52, 152, 219); // 浅蓝色
        private Color accentColor = Color.FromArgb(231, 76, 60);     // 红色
        private Color successColor = Color.FromArgb(46, 204, 113);   // 绿色
        private Color backgroundColor = Color.FromArgb(245, 245, 245); // 浅灰背景

        public MainForm()
        {
            // 检查管理员权限
            if (!IsRunningAsAdministrator())
            {
                ShowAdminWarning();
            }

            InitializeComponent();
            InitializeTrayIcon();
            LoadStageTasks();
        }

        private void InitializeComponent()
        {
            // 主窗体设置
            this.Text = "自动开关机管理器 v3.0";
            this.Size = new Size(650, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Icon = SystemIcons.Shield; // 使用盾牌图标表示需要权限
            this.BackColor = backgroundColor;
            this.Font = new Font("微软雅黑", 9);

            // 初始化TabControl
            InitializeTabControl();

            // 设置Tab键顺序
            SetTabOrder();
        }

        private void InitializeTabControl()
        {
            tabControl = new TabControl();
            tabControl.Location = new Point(15, 15);
            tabControl.Size = new Size(610, 630);
            tabControl.Font = new Font("微软雅黑", 9);
            this.Controls.Add(tabControl);

            // 定时任务标签页
            tabScheduled = new TabPage();
            tabScheduled.Text = "单次定时";
            tabScheduled.Padding = new Padding(15);
            tabScheduled.BackColor = backgroundColor;
            tabControl.Controls.Add(tabScheduled);

            // 立即操作标签页
            tabImmediate = new TabPage();
            tabImmediate.Text = "立即操作";
            tabImmediate.Padding = new Padding(15);
            tabImmediate.BackColor = backgroundColor;
            tabControl.Controls.Add(tabImmediate);

            // 阶段时间标签页
            tabStageTime = new TabPage();
            tabStageTime.Text = "阶段时间";
            tabStageTime.Padding = new Padding(15);
            tabStageTime.BackColor = backgroundColor;
            tabControl.Controls.Add(tabStageTime);

            // 初始化各页面
            InitializeScheduledTab();
            InitializeImmediateTab();
            InitializeStageTimeTab();
        }

        private void InitializeScheduledTab()
        {
            // 初始化开机设置组
            InitializeBootGroup();

            // 初始化关机设置组
            InitializeShutdownGroup();
        }

        private void InitializeBootGroup()
        {
            grpBoot = new GroupBox();
            grpBoot.Text = "单次定时开机";
            grpBoot.Location = new Point(20, 20);
            grpBoot.Size = new Size(550, 160);
            grpBoot.TabIndex = 0;
            grpBoot.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            grpBoot.BackColor = Color.White;
            tabScheduled.Controls.Add(grpBoot);

            // 开机日期标签
            lblBootTime = new Label();
            lblBootTime.Text = "开机时间:";
            lblBootTime.Location = new Point(25, 35);
            lblBootTime.Size = new Size(80, 20);
            lblBootTime.TabIndex = 0;
            grpBoot.Controls.Add(lblBootTime);

            // 开机日期选择器
            dtpBootDate = new DateTimePicker();
            dtpBootDate.Format = DateTimePickerFormat.Short;
            dtpBootDate.Location = new Point(105, 32);
            dtpBootDate.Size = new Size(130, 23);
            dtpBootDate.TabIndex = 1;
            dtpBootDate.Value = DateTime.Today.AddDays(1);
            grpBoot.Controls.Add(dtpBootDate);

            // 开机时间选择器
            dtpBootTime = new DateTimePicker();
            dtpBootTime.Format = DateTimePickerFormat.Custom;
            dtpBootTime.CustomFormat = "HH:mm";
            dtpBootTime.ShowUpDown = true;
            dtpBootTime.Location = new Point(245, 32);
            dtpBootTime.Size = new Size(90, 23);
            dtpBootTime.TabIndex = 2;
            dtpBootTime.Value = DateTime.Today.AddHours(8);
            grpBoot.Controls.Add(dtpBootTime);

            // 启用开机复选框
            chkBootEnabled = new CheckBox();
            chkBootEnabled.Text = "启用定时开机";
            chkBootEnabled.Location = new Point(25, 75);
            chkBootEnabled.Size = new Size(120, 20);
            chkBootEnabled.TabIndex = 3;
            chkBootEnabled.Checked = true;
            grpBoot.Controls.Add(chkBootEnabled);

            // 重复开机复选框
            chkBootRepeat = new CheckBox();
            chkBootRepeat.Text = "每天重复";
            chkBootRepeat.Location = new Point(155, 75);
            chkBootRepeat.Size = new Size(100, 20);
            chkBootRepeat.TabIndex = 4;
            chkBootRepeat.Checked = false;
            grpBoot.Controls.Add(chkBootRepeat);

            // 设置开机任务按钮
            btnSetBoot = CreateStyledButton("设置开机任务", 25, 110, 120, 35);
            btnSetBoot.TabIndex = 5;
            btnSetBoot.Click += SetBootTask_Click;
            grpBoot.Controls.Add(btnSetBoot);

            // 删除开机任务按钮
            btnRemoveBoot = CreateStyledButton("删除开机任务", 155, 110, 120, 35, false);
            btnRemoveBoot.TabIndex = 6;
            btnRemoveBoot.Click += RemoveBootTask_Click;
            grpBoot.Controls.Add(btnRemoveBoot);

            // 开机说明标签
            Label lblBootInfo = new Label();
            lblBootInfo.Text = "设置单次或重复的开机时间，需要主板支持ACPI电源管理";
            lblBootInfo.Location = new Point(300, 75);
            lblBootInfo.Size = new Size(230, 40);
            lblBootInfo.ForeColor = Color.Gray;
            lblBootInfo.Font = new Font("微软雅黑", 8);
            grpBoot.Controls.Add(lblBootInfo);
        }

        private void InitializeShutdownGroup()
        {
            grpShutdown = new GroupBox();
            grpShutdown.Text = "单次定时关机";
            grpShutdown.Location = new Point(20, 200);
            grpShutdown.Size = new Size(550, 160);
            grpShutdown.TabIndex = 1;
            grpShutdown.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            grpShutdown.BackColor = Color.White;
            tabScheduled.Controls.Add(grpShutdown);

            // 关机时间标签
            lblShutdownTime = new Label();
            lblShutdownTime.Text = "关机时间:";
            lblShutdownTime.Location = new Point(25, 35);
            lblShutdownTime.Size = new Size(80, 20);
            lblShutdownTime.TabIndex = 7;
            grpShutdown.Controls.Add(lblShutdownTime);

            // 关机日期选择器
            dtpShutdownDate = new DateTimePicker();
            dtpShutdownDate.Format = DateTimePickerFormat.Short;
            dtpShutdownDate.Location = new Point(105, 32);
            dtpShutdownDate.Size = new Size(130, 23);
            dtpShutdownDate.TabIndex = 8;
            dtpShutdownDate.Value = DateTime.Today;
            grpShutdown.Controls.Add(dtpShutdownDate);

            // 关机时间选择器
            dtpShutdownTime = new DateTimePicker();
            dtpShutdownTime.Format = DateTimePickerFormat.Custom;
            dtpShutdownTime.CustomFormat = "HH:mm";
            dtpShutdownTime.ShowUpDown = true;
            dtpShutdownTime.Location = new Point(245, 32);
            dtpShutdownTime.Size = new Size(90, 23);
            dtpShutdownTime.TabIndex = 9;
            dtpShutdownTime.Value = DateTime.Now.AddHours(1);
            grpShutdown.Controls.Add(dtpShutdownTime);

            // 启机关机复选框
            chkShutdownEnabled = new CheckBox();
            chkShutdownEnabled.Text = "启用定时关机";
            chkShutdownEnabled.Location = new Point(25, 75);
            chkShutdownEnabled.Size = new Size(120, 20);
            chkShutdownEnabled.TabIndex = 10;
            chkShutdownEnabled.Checked = true;
            grpShutdown.Controls.Add(chkShutdownEnabled);

            // 设置关机任务按钮
            btnSetShutdown = CreateStyledButton("设置关机任务", 25, 110, 120, 35);
            btnSetShutdown.TabIndex = 11;
            btnSetShutdown.Click += SetShutdownTask_Click;
            grpShutdown.Controls.Add(btnSetShutdown);

            // 取消关机任务按钮
            btnCancelShutdown = CreateStyledButton("取消关机任务", 155, 110, 120, 35, false);
            btnCancelShutdown.TabIndex = 12;
            btnCancelShutdown.Click += CancelShutdownTask_Click;
            grpShutdown.Controls.Add(btnCancelShutdown);

            // 关机说明标签
            Label lblShutdownInfo = new Label();
            lblShutdownInfo.Text = "设置单次关机时间，关机前请保存所有工作";
            lblShutdownInfo.Location = new Point(300, 75);
            lblShutdownInfo.Size = new Size(230, 40);
            lblShutdownInfo.ForeColor = Color.Gray;
            lblShutdownInfo.Font = new Font("微软雅黑", 8);
            grpShutdown.Controls.Add(lblShutdownInfo);
        }

        private void InitializeImmediateTab()
        {
            // 立即操作组
            grpImmediate = new GroupBox();
            grpImmediate.Text = "立即操作";
            grpImmediate.Location = new Point(50, 30);
            grpImmediate.Size = new Size(500, 220);
            grpImmediate.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            grpImmediate.BackColor = Color.White;
            tabImmediate.Controls.Add(grpImmediate);

            // 立即关机按钮
            btnShutdownNow = CreateStyledButton("立即关机", 50, 50, 120, 45, false, accentColor);
            btnShutdownNow.TabIndex = 0;
            btnShutdownNow.Click += ShutdownNow_Click;
            grpImmediate.Controls.Add(btnShutdownNow);

            // 立即重启按钮
            btnRestartNow = CreateStyledButton("立即重启", 190, 50, 120, 45, false, accentColor);
            btnRestartNow.TabIndex = 1;
            btnRestartNow.Click += RestartNow_Click;
            grpImmediate.Controls.Add(btnRestartNow);

            // 取消关机按钮
            btnCancelNow = CreateStyledButton("取消关机", 330, 50, 120, 45);
            btnCancelNow.TabIndex = 2;
            btnCancelNow.Click += CancelNow_Click;
            grpImmediate.Controls.Add(btnCancelNow);

            // 状态标签
            lblStatus = new Label();
            lblStatus.Text = "当前状态: " + GetShutdownStatus();
            lblStatus.Location = new Point(50, 120);
            lblStatus.Size = new Size(400, 25);
            lblStatus.Name = "lblStatus";
            lblStatus.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;
            UpdateStatusColor();
            grpImmediate.Controls.Add(lblStatus);

            // 权限提示
            Label lblAdminInfo = new Label();
            lblAdminInfo.Text = IsRunningAsAdministrator() ? "✓ 当前以管理员权限运行" : "⚠ 建议以管理员权限运行";
            lblAdminInfo.Location = new Point(50, 160);
            lblAdminInfo.Size = new Size(400, 20);
            lblAdminInfo.ForeColor = IsRunningAsAdministrator() ? successColor : Color.Orange;
            lblAdminInfo.Font = new Font("微软雅黑", 9);
            lblAdminInfo.TextAlign = ContentAlignment.MiddleCenter;
            grpImmediate.Controls.Add(lblAdminInfo);

            // 立即操作说明
            Label lblImmediateInfo = new Label();
            lblImmediateInfo.Text = "警告：立即关机和重启操作会直接执行，请确保已保存所有工作！";
            lblImmediateInfo.Location = new Point(30, 280);
            lblImmediateInfo.Size = new Size(540, 40);
            lblImmediateInfo.ForeColor = accentColor;
            lblImmediateInfo.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            lblImmediateInfo.TextAlign = ContentAlignment.MiddleCenter;
            tabImmediate.Controls.Add(lblImmediateInfo);
        }

        private void InitializeStageTimeTab()
        {
            // 阶段时间设置组
            grpStageSettings = new GroupBox();
            grpStageSettings.Text = "阶段时间设置";
            grpStageSettings.Location = new Point(20, 20);
            grpStageSettings.Size = new Size(550, 280);
            grpStageSettings.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            grpStageSettings.BackColor = Color.White;
            tabStageTime.Controls.Add(grpStageSettings);

            // 星期选择标签
            Label lblDays = new Label();
            lblDays.Text = "选择星期:";
            lblDays.Location = new Point(25, 35);
            lblDays.Size = new Size(80, 20);
            grpStageSettings.Controls.Add(lblDays);

            // 星期选择列表
            chkListDays = new CheckedListBox();
            chkListDays.Location = new Point(105, 35);
            chkListDays.Size = new Size(140, 160);
            chkListDays.Font = new Font("微软雅黑", 9);
            chkListDays.Items.AddRange(new object[] {
                "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"
            });
            // 默认选择周一到周五
            for (int i = 0; i < 5; i++)
            {
                chkListDays.SetItemChecked(i, true);
            }
            grpStageSettings.Controls.Add(chkListDays);

            // 阶段开机时间标签
            lblStageBootTime = new Label();
            lblStageBootTime.Text = "开机时间:";
            lblStageBootTime.Location = new Point(270, 35);
            lblStageBootTime.Size = new Size(80, 20);
            grpStageSettings.Controls.Add(lblStageBootTime);

            // 阶段开机时间选择器
            dtpStageBootTime = new DateTimePicker();
            dtpStageBootTime.Format = DateTimePickerFormat.Custom;
            dtpStageBootTime.CustomFormat = "HH:mm";
            dtpStageBootTime.ShowUpDown = true;
            dtpStageBootTime.Location = new Point(350, 32);
            dtpStageBootTime.Size = new Size(90, 23);
            dtpStageBootTime.Value = DateTime.Today.AddHours(8); // 默认早上8点
            grpStageSettings.Controls.Add(dtpStageBootTime);

            // 阶段关机时间标签
            lblStageShutdownTime = new Label();
            lblStageShutdownTime.Text = "关机时间:";
            lblStageShutdownTime.Location = new Point(270, 75);
            lblStageShutdownTime.Size = new Size(80, 20);
            grpStageSettings.Controls.Add(lblStageShutdownTime);

            // 阶段关机时间选择器
            dtpStageShutdownTime = new DateTimePicker();
            dtpStageShutdownTime.Format = DateTimePickerFormat.Custom;
            dtpStageShutdownTime.CustomFormat = "HH:mm";
            dtpStageShutdownTime.ShowUpDown = true;
            dtpStageShutdownTime.Location = new Point(350, 72);
            dtpStageShutdownTime.Size = new Size(90, 23);
            dtpStageShutdownTime.Value = DateTime.Today.AddHours(18); // 默认晚上6点
            grpStageSettings.Controls.Add(dtpStageShutdownTime);

            // 启用阶段任务复选框
            chkStageEnabled = new CheckBox();
            chkStageEnabled.Text = "启用阶段任务";
            chkStageEnabled.Location = new Point(270, 115);
            chkStageEnabled.Size = new Size(120, 20);
            chkStageEnabled.Checked = true;
            grpStageSettings.Controls.Add(chkStageEnabled);

            // 设置阶段任务按钮
            btnSetStageTask = CreateStyledButton("设置阶段任务", 270, 150, 120, 35);
            btnSetStageTask.Click += SetStageTask_Click;
            grpStageSettings.Controls.Add(btnSetStageTask);

            // 删除阶段任务按钮
            btnRemoveStageTask = CreateStyledButton("删除阶段任务", 400, 150, 120, 35, false);
            btnRemoveStageTask.Click += RemoveStageTask_Click;
            grpStageSettings.Controls.Add(btnRemoveStageTask);

            // 阶段任务说明
            Label lblStageInfo = new Label();
            lblStageInfo.Text = "设置每周固定时间的开关机任务，适合工作日定时开关机";
            lblStageInfo.Location = new Point(25, 210);
            lblStageInfo.Size = new Size(500, 40);
            lblStageInfo.ForeColor = Color.Gray;
            lblStageInfo.Font = new Font("微软雅黑", 8);
            lblStageInfo.TextAlign = ContentAlignment.MiddleCenter;
            grpStageSettings.Controls.Add(lblStageInfo);

            // 阶段任务列表
            listStageTasks = new ListView();
            listStageTasks.Location = new Point(20, 320);
            listStageTasks.Size = new Size(550, 240);
            listStageTasks.View = View.Details;
            listStageTasks.FullRowSelect = true;
            listStageTasks.GridLines = true;
            listStageTasks.Font = new Font("微软雅黑", 9);
            listStageTasks.BackColor = Color.White;

            // 添加列
            listStageTasks.Columns.Add("任务名称", 150);
            listStageTasks.Columns.Add("执行时间", 150);
            listStageTasks.Columns.Add("重复周期", 150);
            listStageTasks.Columns.Add("状态", 80);

            tabStageTime.Controls.Add(listStageTasks);
        }

        // 创建样式化按钮
        private Button CreateStyledButton(string text, int x, int y, int width, int height, bool primary = true, Color? customColor = null)
        {
            Button button = new Button();
            button.Text = text;
            button.Location = new Point(x, y);
            button.Size = new Size(width, height);
            button.FlatStyle = FlatStyle.Flat;
            button.Font = new Font("微软雅黑", 9);

            if (customColor.HasValue)
            {
                button.BackColor = customColor.Value;
                button.ForeColor = Color.White;
            }
            else if (primary)
            {
                button.BackColor = primaryColor;
                button.ForeColor = Color.White;
            }
            else
            {
                button.BackColor = Color.LightGray;
                button.ForeColor = Color.Black;
            }

            button.FlatAppearance.BorderColor = primary ? primaryColor : Color.Gray;
            button.FlatAppearance.MouseOverBackColor = primary ? secondaryColor : Color.Silver;
            button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(31, 97, 141) : Color.DarkGray;

            return button;
        }

        private void SetTabOrder()
        {
            // 设置Tab键顺序
            dtpBootDate.TabIndex = 0;
            dtpBootTime.TabIndex = 1;
            chkBootEnabled.TabIndex = 2;
            chkBootRepeat.TabIndex = 3;
            btnSetBoot.TabIndex = 4;
            btnRemoveBoot.TabIndex = 5;

            dtpShutdownDate.TabIndex = 6;
            dtpShutdownTime.TabIndex = 7;
            chkShutdownEnabled.TabIndex = 8;
            btnSetShutdown.TabIndex = 9;
            btnCancelShutdown.TabIndex = 10;

            btnShutdownNow.TabIndex = 0;
            btnRestartNow.TabIndex = 1;
            btnCancelNow.TabIndex = 2;

            chkListDays.TabIndex = 0;
            dtpStageBootTime.TabIndex = 1;
            dtpStageShutdownTime.TabIndex = 2;
            chkStageEnabled.TabIndex = 3;
            btnSetStageTask.TabIndex = 4;
            btnRemoveStageTask.TabIndex = 5;
        }

        private void InitializeTrayIcon()
        {
            // 创建托盘菜单
            trayMenu = new ContextMenuStrip();
            trayMenu.BackColor = Color.White;
            trayMenu.Font = new Font("微软雅黑", 9);

            menuShow = new ToolStripMenuItem();
            menuShow.Text = "显示窗口";
            menuShow.Click += (s, e) =>
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            };

            menuExit = new ToolStripMenuItem();
            menuExit.Text = "退出";
            menuExit.Click += (s, e) => Application.Exit();

            trayMenu.Items.Add(menuShow);
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(menuExit);

            // 创建托盘图标
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Shield;
            trayIcon.Text = "自动开关机管理器";
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
            trayIcon.DoubleClick += (s, e) =>
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            };

            // 窗体最小化时隐藏到托盘
            this.Resize += (s, e) =>
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.Hide();
                    trayIcon.ShowBalloonTip(1000, "自动开关机管理器", "程序已最小化到系统托盘", ToolTipIcon.Info);
                }
            };
        }

        // 事件处理方法
        private void SetBootTask_Click(object sender, EventArgs e)
        {
            DateTime bootDateTime = dtpBootDate.Value.Date + dtpBootTime.Value.TimeOfDay;
            SetBootTask(bootDateTime, chkBootEnabled.Checked, chkBootRepeat.Checked);
        }

        private void RemoveBootTask_Click(object sender, EventArgs e)
        {
            RemoveBootTask();
        }

        private void SetShutdownTask_Click(object sender, EventArgs e)
        {
            DateTime shutdownDateTime = dtpShutdownDate.Value.Date + dtpShutdownTime.Value.TimeOfDay;
            SetShutdownTask(shutdownDateTime, chkShutdownEnabled.Checked);
        }

        private void CancelShutdownTask_Click(object sender, EventArgs e)
        {
            CancelShutdownTask();
            UpdateStatusLabel();
        }

        private void ShutdownNow_Click(object sender, EventArgs e)
        {
            ShutdownNow();
        }

        private void RestartNow_Click(object sender, EventArgs e)
        {
            RestartNow();
        }

        private void CancelNow_Click(object sender, EventArgs e)
        {
            CancelNow();
            UpdateStatusLabel();
        }

        private void SetStageTask_Click(object sender, EventArgs e)
        {
            SetStageTask();
        }

        private void RemoveStageTask_Click(object sender, EventArgs e)
        {
            RemoveStageTask();
        }

        // 功能实现方法
        private void SetBootTask(DateTime bootTime, bool enabled, bool repeat)
        {
            try
            {
                if (bootTime <= DateTime.Now && !repeat)
                {
                    MessageBox.Show("开机时间必须是未来的时间", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (TaskService taskService = new TaskService())
                {
                    // 删除已存在的任务
                    try
                    {
                        taskService.RootFolder.DeleteTask("AutoPowerOn", false);
                    }
                    catch
                    {
                        // 任务不存在，忽略错误
                    }

                    if (enabled)
                    {
                        // 创建开机任务
                        TaskDefinition bootTask = taskService.NewTask();
                        bootTask.RegistrationInfo.Description = "自动开机任务";
                        bootTask.Principal.RunLevel = TaskRunLevel.Highest;

                        if (repeat)
                        {
                            // 每天重复
                            bootTask.Triggers.Add(new DailyTrigger
                            {
                                StartBoundary = bootTime,
                                Enabled = true
                            });
                        }
                        else
                        {
                            // 单次执行
                            bootTask.Triggers.Add(new TimeTrigger
                            {
                                StartBoundary = bootTime,
                                Enabled = true
                            });
                        }

                        // 设置操作 - 使用计划任务唤醒计算机
                        bootTask.Actions.Add(new ExecAction("cmd.exe", "/c echo Wake up task", null));
                        bootTask.Settings.WakeToRun = true;
                        bootTask.Settings.DisallowStartIfOnBatteries = false;
                        bootTask.Settings.StopIfGoingOnBatteries = false;

                        taskService.RootFolder.RegisterTaskDefinition("AutoPowerOn", bootTask);

                        string repeatText = repeat ? "每天" : "一次";
                        MessageBox.Show($"定时开机已设置为 {bootTime:yyyy-MM-dd HH:mm} ({repeatText})", "成功",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("开机任务已禁用", "信息",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveBootTask()
        {
            try
            {
                using (TaskService taskService = new TaskService())
                {
                    taskService.RootFolder.DeleteTask("AutoPowerOn", false);
                    MessageBox.Show("开机任务已删除", "成功",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除开机任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetShutdownTask(DateTime shutdownTime, bool enabled)
        {
            try
            {
                if (shutdownTime <= DateTime.Now)
                {
                    MessageBox.Show("关机时间必须是未来的时间", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                TimeSpan timeUntilShutdown = shutdownTime - DateTime.Now;
                int seconds = (int)timeUntilShutdown.TotalSeconds;

                if (seconds <= 0)
                {
                    MessageBox.Show("关机时间必须是将来的时间", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (enabled)
                {
                    // 使用shutdown命令设置定时关机
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "shutdown",
                        Arguments = $"/s /t {seconds}",
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });

                    MessageBox.Show($"系统将在 {shutdownTime:yyyy-MM-dd HH:mm} 关机", "关机任务已设置",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    CancelShutdownTask();
                }

                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置关机任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CancelShutdownTask()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "shutdown",
                    Arguments = "/a",
                    UseShellExecute = true,
                    Verb = IsRunningAsAdministrator() ? "runas" : ""
                });

                MessageBox.Show("关机任务已取消", "成功",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取消关机任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShutdownNow()
        {
            if (MessageBox.Show("确定要立即关机吗？请确保已保存所有工作！", "确认关机",
                              MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "shutdown",
                        Arguments = "/s /t 0",
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"关机失败: {ex.Message}", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void RestartNow()
        {
            if (MessageBox.Show("确定要立即重启吗？请确保已保存所有工作！", "确认重启",
                              MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "shutdown",
                        Arguments = "/r /t 0",
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"重启失败: {ex.Message}", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void CancelNow()
        {
            CancelShutdownTask();
        }

        private void SetStageTask()
        {
            try
            {
                // 检查是否选择了星期
                if (chkListDays.CheckedIndices.Count == 0)
                {
                    MessageBox.Show("请至少选择一个星期", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查时间是否合理
                if (dtpStageBootTime.Value.TimeOfDay >= dtpStageShutdownTime.Value.TimeOfDay)
                {
                    MessageBox.Show("关机时间必须晚于开机时间", "错误",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (TaskService taskService = new TaskService())
                {
                    // 删除已存在的阶段任务
                    try
                    {
                        taskService.RootFolder.DeleteTask("StagePowerOn", false);
                        taskService.RootFolder.DeleteTask("StagePowerOff", false);
                    }
                    catch
                    {
                        // 任务不存在，忽略错误
                    }

                    if (chkStageEnabled.Checked)
                    {
                        // 创建阶段开机任务
                        TaskDefinition stageBootTask = taskService.NewTask();
                        stageBootTask.RegistrationInfo.Description = "阶段开机任务";
                        stageBootTask.Principal.RunLevel = TaskRunLevel.Highest;

                        // 为每个选中的星期创建触发器
                        foreach (int dayIndex in chkListDays.CheckedIndices)
                        {
                            // 计算下一个该星期几的日期
                            DateTime nextDate = GetNextWeekday((DayOfWeek)(dayIndex + 1));
                            DateTime bootTime = nextDate.Date + dtpStageBootTime.Value.TimeOfDay;

                            stageBootTask.Triggers.Add(new WeeklyTrigger
                            {
                                DaysOfWeek = (DaysOfTheWeek)(1 << dayIndex),
                                StartBoundary = bootTime,
                                Enabled = true
                            });
                        }

                        // 设置操作
                        stageBootTask.Actions.Add(new ExecAction("cmd.exe", "/c echo Stage boot task", null));
                        stageBootTask.Settings.WakeToRun = true;
                        stageBootTask.Settings.DisallowStartIfOnBatteries = false;
                        stageBootTask.Settings.StopIfGoingOnBatteries = false;

                        taskService.RootFolder.RegisterTaskDefinition("StagePowerOn", stageBootTask);

                        // 创建阶段关机任务
                        TaskDefinition stageShutdownTask = taskService.NewTask();
                        stageShutdownTask.RegistrationInfo.Description = "阶段关机任务";
                        stageShutdownTask.Principal.RunLevel = TaskRunLevel.Highest;

                        // 为每个选中的星期创建触发器
                        foreach (int dayIndex in chkListDays.CheckedIndices)
                        {
                            // 计算下一个该星期几的日期
                            DateTime nextDate = GetNextWeekday((DayOfWeek)(dayIndex + 1));
                            DateTime shutdownTime = nextDate.Date + dtpStageShutdownTime.Value.TimeOfDay;

                            stageShutdownTask.Triggers.Add(new WeeklyTrigger
                            {
                                DaysOfWeek = (DaysOfTheWeek)(1 << dayIndex),
                                StartBoundary = shutdownTime,
                                Enabled = true
                            });
                        }

                        // 设置关机操作
                        stageShutdownTask.Actions.Add(new ExecAction("shutdown", "/s /t 0", null));

                        taskService.RootFolder.RegisterTaskDefinition("StagePowerOff", stageShutdownTask);

                        // 获取选择的星期文本
                        string daysText = GetSelectedDaysText();

                        MessageBox.Show($"阶段任务已设置:\n开机: {dtpStageBootTime.Value:HH:mm}\n关机: {dtpStageShutdownTime.Value:HH:mm}\n执行: {daysText}",
                                      "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 刷新任务列表
                        LoadStageTasks();
                    }
                    else
                    {
                        MessageBox.Show("阶段任务已禁用", "信息",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置阶段任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveStageTask()
        {
            try
            {
                using (TaskService taskService = new TaskService())
                {
                    taskService.RootFolder.DeleteTask("StagePowerOn", false);
                    taskService.RootFolder.DeleteTask("StagePowerOff", false);
                    MessageBox.Show("阶段任务已删除", "成功",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 刷新任务列表
                    LoadStageTasks();
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除阶段任务失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadStageTasks()
        {
            listStageTasks.Items.Clear();

            try
            {
                using (TaskService taskService = new TaskService())
                {
                    // 检查阶段开机任务
                    Task stageBootTask = taskService.FindTask("StagePowerOn");
                    if (stageBootTask != null)
                    {
                        ListViewItem item = new ListViewItem("阶段开机");
                        item.SubItems.Add(dtpStageBootTime.Value.ToString("HH:mm"));
                        item.SubItems.Add(GetSelectedDaysText());
                        item.SubItems.Add("已启用");
                        listStageTasks.Items.Add(item);
                    }

                    // 检查阶段关机任务
                    Task stageShutdownTask = taskService.FindTask("StagePowerOff");
                    if (stageShutdownTask != null)
                    {
                        ListViewItem item = new ListViewItem("阶段关机");
                        item.SubItems.Add(dtpStageShutdownTime.Value.ToString("HH:mm"));
                        item.SubItems.Add(GetSelectedDaysText());
                        item.SubItems.Add("已启用");
                        listStageTasks.Items.Add(item);
                    }

                    // 检查单次开机任务
                    Task bootTask = taskService.FindTask("AutoPowerOn");
                    if (bootTask != null && bootTask.Enabled)
                    {
                        foreach (Trigger trigger in bootTask.Definition.Triggers)
                        {
                            ListViewItem item = new ListViewItem("单次开机");
                            item.SubItems.Add(trigger.StartBoundary.ToString("yyyy-MM-dd HH:mm"));
                            item.SubItems.Add(trigger is DailyTrigger ? "每天" : "一次");
                            item.SubItems.Add("已启用");
                            listStageTasks.Items.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 忽略加载任务列表时的错误
            }
        }

        // 辅助方法
        private DateTime GetNextWeekday(DayOfWeek day)
        {
            DateTime result = DateTime.Now.AddDays(1);
            while (result.DayOfWeek != day)
                result = result.AddDays(1);
            return result;
        }

        private string GetSelectedDaysText()
        {
            List<string> selectedDays = new List<string>();
            for (int i = 0; i < chkListDays.Items.Count; i++)
            {
                if (chkListDays.GetItemChecked(i))
                {
                    selectedDays.Add(chkListDays.Items[i].ToString());
                }
            }
            return string.Join(", ", selectedDays);
        }

        private string GetShutdownStatus()
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "shutdown";
                process.StartInfo.Arguments = "/a"; // 尝试取消关机来检测是否有计划
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                // 如果有计划关机，会显示取消消息
                if (output.Contains("注销被取消"))
                {
                    return "有计划的关机任务";
                }
            }
            catch
            {
                // 忽略错误
            }

            return "无计划关机任务";
        }

        private void UpdateStatusLabel()
        {
            lblStatus.Text = "当前状态: " + GetShutdownStatus();
            UpdateStatusColor();
        }

        private void UpdateStatusColor()
        {
            if (lblStatus.Text.Contains("有计划的关机任务"))
            {
                lblStatus.ForeColor = accentColor;
            }
            else
            {
                lblStatus.ForeColor = successColor;
            }
        }

        // 权限检查方法
        private bool IsRunningAsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void ShowAdminWarning()
        {
            MessageBox.Show("检测到当前未以管理员权限运行。\n\n部分功能（如定时开机）需要管理员权限才能正常工作。\n\n建议右键点击程序，选择\"以管理员身份运行\"。",
                          "权限提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowAdminRequiredMessage()
        {
            MessageBox.Show("操作需要管理员权限！\n\n请右键点击程序，选择\"以管理员身份运行\"，然后重试。",
                          "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                trayIcon.ShowBalloonTip(1000, "自动开关机管理器", "程序已最小化到系统托盘", ToolTipIcon.Info);
            }
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                trayIcon?.Dispose();
                trayMenu?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    internal class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 检查是否以管理员身份运行
            if (!IsRunningAsAdministrator())
            {
                DialogResult result = MessageBox.Show("检测到当前未以管理员权限运行。\n\n部分功能需要管理员权限才能正常工作。\n\n是否尝试以管理员权限重新启动？",
                                                    "权限提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // 重新启动并请求管理员权限
                    RestartAsAdmin();
                    return;
                }
            }

            Application.Run(new MainForm());
        }

        static bool IsRunningAsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RestartAsAdmin()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = true;
            startInfo.WorkingDirectory = Environment.CurrentDirectory;
            startInfo.FileName = Application.ExecutablePath;
            startInfo.Verb = "runas"; // 请求管理员权限

            try
            {
                Process.Start(startInfo);
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法以管理员权限重新启动: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}