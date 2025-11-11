using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Security.Principal;
using System.IO;
using Microsoft.Win32;
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
        private TabPage tabSettings;

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

        // 设置控件
        private GroupBox grpSettings;
        private CheckBox chkAutoStart;
        private CheckBox chkStartMinimized;
        private CheckBox chkShowNotifications;
        private Button btnSaveSettings;
        private Label lblVersion;

        // 系统托盘
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem menuShow;
        private ToolStripMenuItem menuExit;

        // 颜色方案 - 现代化设计
        private Color primaryColor = Color.FromArgb(0, 120, 215);    // 现代蓝色
        private Color secondaryColor = Color.FromArgb(23, 147, 209); // 浅蓝色
        private Color accentColor = Color.FromArgb(216, 59, 1);      // 橙色强调色
        private Color successColor = Color.FromArgb(16, 137, 62);    // 深绿色
        private Color warningColor = Color.FromArgb(255, 185, 0);    // 警告黄色
        private Color backgroundColor = Color.FromArgb(250, 250, 250); // 更浅的背景
        private Color cardColor = Color.White;                       // 卡片背景
        private Color borderColor = Color.FromArgb(225, 225, 225);   // 边框颜色
        private Color textColor = Color.FromArgb(68, 68, 68);        // 主要文字颜色
        private Color mutedTextColor = Color.FromArgb(102, 102, 102); // 次要文字颜色

        // 设置存储
        private bool autoStart = false;
        private bool startMinimized = false;
        private bool showNotifications = true;

        // 任务名称常量
        private const string TASK_BOOT = "AutoPowerManager_Boot";
        private const string TASK_SHUTDOWN = "AutoPowerManager_Shutdown";
        private const string TASK_STAGE_BOOT = "AutoPowerManager_StageBoot";
        private const string TASK_STAGE_SHUTDOWN = "AutoPowerManager_StageShutdown";

        public MainForm()
        {
            // 检查管理员权限
            if (!IsRunningAsAdministrator())
            {
                ShowAdminWarning();
            }

            LoadSettings();
            InitializeComponent();
            InitializeTrayIcon();
            LoadStageTasks();
            ApplySettings();
            UpdateStatusLabel();
        }

        private void InitializeComponent()
        {
            // 主窗体设置
            this.Text = "自动开关机管理器 v3.1";
            this.Size = new Size(700, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Icon = SystemIcons.Shield;
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
            tabControl.Size = new Size(660, 680);
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

            // 设置标签页
            tabSettings = new TabPage();
            tabSettings.Text = "设置";
            tabSettings.Padding = new Padding(15);
            tabSettings.BackColor = backgroundColor;
            tabControl.Controls.Add(tabSettings);

            // 初始化各页面
            InitializeScheduledTab();
            InitializeImmediateTab();
            InitializeStageTimeTab();
            InitializeSettingsTab();
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
            grpBoot.Size = new Size(600, 180);
            grpBoot.TabIndex = 0;
            grpBoot.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            grpBoot.ForeColor = textColor;
            grpBoot.BackColor = cardColor;
            grpBoot.Padding = new Padding(20);
            tabScheduled.Controls.Add(grpBoot);

            // 开机日期标签
            lblBootTime = new Label();
            lblBootTime.Text = "开机时间:";
            lblBootTime.Location = new Point(25, 35);
            lblBootTime.Size = new Size(80, 25);
            lblBootTime.TabIndex = 0;
            lblBootTime.Font = new Font("微软雅黑", 9);
            lblBootTime.ForeColor = textColor;
            grpBoot.Controls.Add(lblBootTime);

            // 开机日期选择器
            dtpBootDate = new DateTimePicker();
            dtpBootDate.Format = DateTimePickerFormat.Short;
            dtpBootDate.Location = new Point(110, 32);
            dtpBootDate.Size = new Size(140, 25);
            dtpBootDate.TabIndex = 1;
            dtpBootDate.Value = DateTime.Today.AddDays(1);
            dtpBootDate.Font = new Font("微软雅黑", 9);
            grpBoot.Controls.Add(dtpBootDate);

            // 开机时间选择器
            dtpBootTime = new DateTimePicker();
            dtpBootTime.Format = DateTimePickerFormat.Custom;
            dtpBootTime.CustomFormat = "HH:mm";
            dtpBootTime.ShowUpDown = true;
            dtpBootTime.Location = new Point(260, 32);
            dtpBootTime.Size = new Size(100, 25);
            dtpBootTime.TabIndex = 2;
            dtpBootTime.Value = DateTime.Today.AddHours(8);
            dtpBootTime.Font = new Font("微软雅黑", 9);
            grpBoot.Controls.Add(dtpBootTime);

            // 启用开机复选框
            chkBootEnabled = new CheckBox();
            chkBootEnabled.Text = "启用定时开机";
            chkBootEnabled.Location = new Point(25, 80);
            chkBootEnabled.Size = new Size(120, 25);
            chkBootEnabled.TabIndex = 3;
            chkBootEnabled.Checked = true;
            chkBootEnabled.Font = new Font("微软雅黑", 9);
            chkBootEnabled.ForeColor = textColor;
            grpBoot.Controls.Add(chkBootEnabled);

            // 重复开机复选框
            chkBootRepeat = new CheckBox();
            chkBootRepeat.Text = "每天重复";
            chkBootRepeat.Location = new Point(160, 80);
            chkBootRepeat.Size = new Size(100, 25);
            chkBootRepeat.TabIndex = 4;
            chkBootRepeat.Checked = false;
            chkBootRepeat.Font = new Font("微软雅黑", 9);
            chkBootRepeat.ForeColor = textColor;
            grpBoot.Controls.Add(chkBootRepeat);

            // 设置开机任务按钮
            btnSetBoot = CreateStyledButton("设置开机任务", 25, 120, 130, 40);
            btnSetBoot.TabIndex = 5;
            btnSetBoot.Click += SetBootTask_Click;
            grpBoot.Controls.Add(btnSetBoot);

            // 删除开机任务按钮
            btnRemoveBoot = CreateStyledButton("删除开机任务", 170, 120, 130, 40, false);
            btnRemoveBoot.TabIndex = 6;
            btnRemoveBoot.Click += RemoveBootTask_Click;
            grpBoot.Controls.Add(btnRemoveBoot);

            // 开机说明标签
            Label lblBootInfo = new Label();
            lblBootInfo.Text = "注意：定时开机需要主板支持ACPI电源管理，且仅在接通电源时有效";
            lblBootInfo.Location = new Point(320, 80);
            lblBootInfo.Size = new Size(250, 40);
            lblBootInfo.ForeColor = mutedTextColor;
            lblBootInfo.Font = new Font("微软雅黑", 8);
            lblBootInfo.TextAlign = ContentAlignment.MiddleLeft;
            grpBoot.Controls.Add(lblBootInfo);
        }

        private void InitializeShutdownGroup()
        {
            grpShutdown = new GroupBox();
            grpShutdown.Text = "单次定时关机";
            grpShutdown.Location = new Point(20, 220);
            grpShutdown.Size = new Size(600, 180);
            grpShutdown.TabIndex = 1;
            grpShutdown.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            grpShutdown.ForeColor = textColor;
            grpShutdown.BackColor = cardColor;
            grpShutdown.Padding = new Padding(20);
            tabScheduled.Controls.Add(grpShutdown);

            // 关机时间标签
            lblShutdownTime = new Label();
            lblShutdownTime.Text = "关机时间:";
            lblShutdownTime.Location = new Point(25, 35);
            lblShutdownTime.Size = new Size(80, 25);
            lblShutdownTime.TabIndex = 7;
            lblShutdownTime.Font = new Font("微软雅黑", 9);
            lblShutdownTime.ForeColor = textColor;
            grpShutdown.Controls.Add(lblShutdownTime);

            // 关机日期选择器
            dtpShutdownDate = new DateTimePicker();
            dtpShutdownDate.Format = DateTimePickerFormat.Short;
            dtpShutdownDate.Location = new Point(110, 32);
            dtpShutdownDate.Size = new Size(140, 25);
            dtpShutdownDate.TabIndex = 8;
            dtpShutdownDate.Value = DateTime.Today;
            dtpShutdownDate.Font = new Font("微软雅黑", 9);
            grpShutdown.Controls.Add(dtpShutdownDate);

            // 关机时间选择器
            dtpShutdownTime = new DateTimePicker();
            dtpShutdownTime.Format = DateTimePickerFormat.Custom;
            dtpShutdownTime.CustomFormat = "HH:mm";
            dtpShutdownTime.ShowUpDown = true;
            dtpShutdownTime.Location = new Point(260, 32);
            dtpShutdownTime.Size = new Size(100, 25);
            dtpShutdownTime.TabIndex = 9;
            dtpShutdownTime.Value = DateTime.Now.AddHours(1);
            dtpShutdownTime.Font = new Font("微软雅黑", 9);
            grpShutdown.Controls.Add(dtpShutdownTime);

            // 启机关机复选框
            chkShutdownEnabled = new CheckBox();
            chkShutdownEnabled.Text = "启用定时关机";
            chkShutdownEnabled.Location = new Point(25, 80);
            chkShutdownEnabled.Size = new Size(120, 25);
            chkShutdownEnabled.TabIndex = 10;
            chkShutdownEnabled.Checked = true;
            chkShutdownEnabled.Font = new Font("微软雅黑", 9);
            chkShutdownEnabled.ForeColor = textColor;
            grpShutdown.Controls.Add(chkShutdownEnabled);

            // 设置关机任务按钮
            btnSetShutdown = CreateStyledButton("设置关机任务", 25, 120, 130, 40);
            btnSetShutdown.TabIndex = 11;
            btnSetShutdown.Click += SetShutdownTask_Click;
            grpShutdown.Controls.Add(btnSetShutdown);

            // 取消关机任务按钮
            btnCancelShutdown = CreateStyledButton("取消关机任务", 170, 120, 130, 40, false);
            btnCancelShutdown.TabIndex = 12;
            btnCancelShutdown.Click += CancelShutdownTask_Click;
            grpShutdown.Controls.Add(btnCancelShutdown);

            // 关机说明标签
            Label lblShutdownInfo = new Label();
            lblShutdownInfo.Text = "设置单次关机时间，关机前请保存所有工作。关机前1分钟会有提醒";
            lblShutdownInfo.Location = new Point(320, 80);
            lblShutdownInfo.Size = new Size(250, 40);
            lblShutdownInfo.ForeColor = mutedTextColor;
            lblShutdownInfo.Font = new Font("微软雅黑", 8);
            lblShutdownInfo.TextAlign = ContentAlignment.MiddleLeft;
            grpShutdown.Controls.Add(lblShutdownInfo);
        }

        private void InitializeImmediateTab()
        {
            // 立即操作组
            grpImmediate = new GroupBox();
            grpImmediate.Text = "立即操作";
            grpImmediate.Location = new Point(30, 30);
            grpImmediate.Size = new Size(600, 280);
            grpImmediate.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            grpImmediate.BackColor = cardColor;
            grpImmediate.ForeColor = textColor;
            grpImmediate.Padding = new Padding(20);
            tabImmediate.Controls.Add(grpImmediate);

            // 按钮容器面板
            Panel buttonPanel = new Panel();
            buttonPanel.Location = new Point(50, 40);
            buttonPanel.Size = new Size(500, 80);
            buttonPanel.BackColor = Color.Transparent;
            grpImmediate.Controls.Add(buttonPanel);

            // 立即关机按钮
            btnShutdownNow = CreateStyledButton("立即关机", 0, 0, 140, 50, false, accentColor);
            btnShutdownNow.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            btnShutdownNow.TabIndex = 0;
            btnShutdownNow.Click += ShutdownNow_Click;
            buttonPanel.Controls.Add(btnShutdownNow);

            // 立即重启按钮
            btnRestartNow = CreateStyledButton("立即重启", 180, 0, 140, 50, false, accentColor);
            btnRestartNow.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            btnRestartNow.TabIndex = 1;
            btnRestartNow.Click += RestartNow_Click;
            buttonPanel.Controls.Add(btnRestartNow);

            // 取消关机按钮
            btnCancelNow = CreateStyledButton("取消关机", 360, 0, 140, 50);
            btnCancelNow.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            btnCancelNow.TabIndex = 2;
            btnCancelNow.Click += CancelNow_Click;
            buttonPanel.Controls.Add(btnCancelNow);

            // 状态标签容器
            Panel statusPanel = new Panel();
            statusPanel.Location = new Point(50, 140);
            statusPanel.Size = new Size(500, 60);
            statusPanel.BackColor = Color.FromArgb(245, 245, 245);
            statusPanel.BorderStyle = BorderStyle.FixedSingle;
            statusPanel.Padding = new Padding(10);
            grpImmediate.Controls.Add(statusPanel);

            // 状态标签
            lblStatus = new Label();
            lblStatus.Text = "当前状态: " + GetShutdownStatus();
            lblStatus.Location = new Point(0, 15);
            lblStatus.Size = new Size(480, 30);
            lblStatus.Name = "lblStatus";
            lblStatus.Font = new Font("微软雅黑", 11, FontStyle.Bold);
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;
            UpdateStatusColor();
            statusPanel.Controls.Add(lblStatus);

            // 权限提示容器
            Panel adminPanel = new Panel();
            adminPanel.Location = new Point(50, 220);
            adminPanel.Size = new Size(500, 30);
            adminPanel.BackColor = Color.Transparent;
            grpImmediate.Controls.Add(adminPanel);

            // 权限提示
            Label lblAdminInfo = new Label();
            lblAdminInfo.Text = IsRunningAsAdministrator() ? "✓ 当前以管理员权限运行" : "⚠ 建议以管理员权限运行";
            lblAdminInfo.Location = new Point(0, 5);
            lblAdminInfo.Size = new Size(500, 20);
            lblAdminInfo.ForeColor = IsRunningAsAdministrator() ? successColor : warningColor;
            lblAdminInfo.Font = new Font("微软雅黑", 9, FontStyle.Bold);
            lblAdminInfo.TextAlign = ContentAlignment.MiddleCenter;
            adminPanel.Controls.Add(lblAdminInfo);

            // 立即操作说明
            Label lblImmediateInfo = new Label();
            lblImmediateInfo.Text = "警告：立即关机和重启操作会直接执行，请确保已保存所有工作！";
            lblImmediateInfo.Location = new Point(30, 330);
            lblImmediateInfo.Size = new Size(600, 40);
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
            grpStageSettings.Size = new Size(600, 300);
            grpStageSettings.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            grpStageSettings.ForeColor = textColor;
            grpStageSettings.BackColor = cardColor;
            grpStageSettings.Padding = new Padding(20);
            tabStageTime.Controls.Add(grpStageSettings);

            // 星期选择标签
            Label lblDays = new Label();
            lblDays.Text = "选择星期:";
            lblDays.Location = new Point(25, 35);
            lblDays.Size = new Size(80, 25);
            lblDays.Font = new Font("微软雅黑", 9);
            lblDays.ForeColor = textColor;
            grpStageSettings.Controls.Add(lblDays);

            // 星期选择列表
            chkListDays = new CheckedListBox();
            chkListDays.Location = new Point(110, 35);
            chkListDays.Size = new Size(150, 180);
            chkListDays.Font = new Font("微软雅黑", 9);
            chkListDays.BackColor = Color.White;
            chkListDays.BorderStyle = BorderStyle.FixedSingle;
            chkListDays.Items.AddRange(new object[] {
                "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"
            });
            // 默认选择周一到周五
            for (int i = 0; i < 5; i++)
            {
                chkListDays.SetItemChecked(i, true);
            }
            grpStageSettings.Controls.Add(chkListDays);

            // 时间设置面板
            Panel timePanel = new Panel();
            timePanel.Location = new Point(280, 35);
            timePanel.Size = new Size(280, 180);
            timePanel.BackColor = Color.Transparent;
            grpStageSettings.Controls.Add(timePanel);

            // 阶段开机时间标签
            lblStageBootTime = new Label();
            lblStageBootTime.Text = "开机时间:";
            lblStageBootTime.Location = new Point(0, 10);
            lblStageBootTime.Size = new Size(80, 25);
            lblStageBootTime.Font = new Font("微软雅黑", 9);
            lblStageBootTime.ForeColor = textColor;
            timePanel.Controls.Add(lblStageBootTime);

            // 阶段开机时间选择器
            dtpStageBootTime = new DateTimePicker();
            dtpStageBootTime.Format = DateTimePickerFormat.Custom;
            dtpStageBootTime.CustomFormat = "HH:mm";
            dtpStageBootTime.ShowUpDown = true;
            dtpStageBootTime.Location = new Point(85, 7);
            dtpStageBootTime.Size = new Size(100, 25);
            dtpStageBootTime.Value = DateTime.Today.AddHours(8); // 默认早上8点
            dtpStageBootTime.Font = new Font("微软雅黑", 9);
            timePanel.Controls.Add(dtpStageBootTime);

            // 阶段关机时间标签
            lblStageShutdownTime = new Label();
            lblStageShutdownTime.Text = "关机时间:";
            lblStageShutdownTime.Location = new Point(0, 50);
            lblStageShutdownTime.Size = new Size(80, 25);
            lblStageShutdownTime.Font = new Font("微软雅黑", 9);
            lblStageShutdownTime.ForeColor = textColor;
            timePanel.Controls.Add(lblStageShutdownTime);

            // 阶段关机时间选择器
            dtpStageShutdownTime = new DateTimePicker();
            dtpStageShutdownTime.Format = DateTimePickerFormat.Custom;
            dtpStageShutdownTime.CustomFormat = "HH:mm";
            dtpStageShutdownTime.ShowUpDown = true;
            dtpStageShutdownTime.Location = new Point(85, 47);
            dtpStageShutdownTime.Size = new Size(100, 25);
            dtpStageShutdownTime.Value = DateTime.Today.AddHours(18); // 默认晚上6点
            dtpStageShutdownTime.Font = new Font("微软雅黑", 9);
            timePanel.Controls.Add(dtpStageShutdownTime);

            // 启用阶段任务复选框
            chkStageEnabled = new CheckBox();
            chkStageEnabled.Text = "启用阶段任务";
            chkStageEnabled.Location = new Point(0, 90);
            chkStageEnabled.Size = new Size(120, 25);
            chkStageEnabled.Checked = true;
            chkStageEnabled.Font = new Font("微软雅黑", 9);
            chkStageEnabled.ForeColor = textColor;
            timePanel.Controls.Add(chkStageEnabled);

            // 按钮容器
            Panel buttonPanel = new Panel();
            buttonPanel.Location = new Point(0, 125);
            buttonPanel.Size = new Size(280, 50);
            buttonPanel.BackColor = Color.Transparent;
            timePanel.Controls.Add(buttonPanel);

            // 设置阶段任务按钮
            btnSetStageTask = CreateStyledButton("设置阶段任务", 0, 0, 130, 40);
            btnSetStageTask.Click += SetStageTask_Click;
            buttonPanel.Controls.Add(btnSetStageTask);

            // 删除阶段任务按钮
            btnRemoveStageTask = CreateStyledButton("删除阶段任务", 140, 0, 130, 40, false);
            btnRemoveStageTask.Click += RemoveStageTask_Click;
            buttonPanel.Controls.Add(btnRemoveStageTask);

            // 阶段任务说明
            Label lblStageInfo = new Label();
            lblStageInfo.Text = "设置每周固定时间的开关机任务，适合工作日定时开关机";
            lblStageInfo.Location = new Point(25, 240);
            lblStageInfo.Size = new Size(550, 40);
            lblStageInfo.ForeColor = mutedTextColor;
            lblStageInfo.Font = new Font("微软雅黑", 8);
            lblStageInfo.TextAlign = ContentAlignment.MiddleCenter;
            grpStageSettings.Controls.Add(lblStageInfo);

            // 阶段任务列表
            listStageTasks = new ListView();
            listStageTasks.Location = new Point(20, 340);
            listStageTasks.Size = new Size(600, 220);
            listStageTasks.View = View.Details;
            listStageTasks.FullRowSelect = true;
            listStageTasks.GridLines = true;
            listStageTasks.Font = new Font("微软雅黑", 9);
            listStageTasks.BackColor = Color.White;
            listStageTasks.BorderStyle = BorderStyle.FixedSingle;

            // 添加列
            listStageTasks.Columns.Add("任务名称", 150);
            listStageTasks.Columns.Add("执行时间", 150);
            listStageTasks.Columns.Add("重复周期", 150);
            listStageTasks.Columns.Add("状态", 80);

            tabStageTime.Controls.Add(listStageTasks);
        }

        private void InitializeSettingsTab()
        {
            // 设置组
            grpSettings = new GroupBox();
            grpSettings.Text = "程序设置";
            grpSettings.Location = new Point(20, 20);
            grpSettings.Size = new Size(600, 280);
            grpSettings.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            grpSettings.ForeColor = textColor;
            grpSettings.BackColor = cardColor;
            grpSettings.Padding = new Padding(20);
            tabSettings.Controls.Add(grpSettings);

            // 设置选项面板
            Panel settingsPanel = new Panel();
            settingsPanel.Location = new Point(25, 35);
            settingsPanel.Size = new Size(550, 120);
            settingsPanel.BackColor = Color.Transparent;
            grpSettings.Controls.Add(settingsPanel);

            // 开机自启动复选框
            chkAutoStart = new CheckBox();
            chkAutoStart.Text = "开机自动启动程序";
            chkAutoStart.Location = new Point(0, 0);
            chkAutoStart.Size = new Size(250, 30);
            chkAutoStart.Checked = autoStart;
            chkAutoStart.Font = new Font("微软雅黑", 9);
            chkAutoStart.ForeColor = textColor;
            chkAutoStart.CheckedChanged += (s, e) => UpdateAutoStart();
            settingsPanel.Controls.Add(chkAutoStart);

            // 启动时最小化复选框
            chkStartMinimized = new CheckBox();
            chkStartMinimized.Text = "启动时最小化到系统托盘";
            chkStartMinimized.Location = new Point(0, 40);
            chkStartMinimized.Size = new Size(280, 30);
            chkStartMinimized.Checked = startMinimized;
            chkStartMinimized.Font = new Font("微软雅黑", 9);
            chkStartMinimized.ForeColor = textColor;
            settingsPanel.Controls.Add(chkStartMinimized);

            // 显示通知复选框
            chkShowNotifications = new CheckBox();
            chkShowNotifications.Text = "显示系统托盘通知";
            chkShowNotifications.Location = new Point(0, 80);
            chkShowNotifications.Size = new Size(250, 30);
            chkShowNotifications.Checked = showNotifications;
            chkShowNotifications.Font = new Font("微软雅黑", 9);
            chkShowNotifications.ForeColor = textColor;
            settingsPanel.Controls.Add(chkShowNotifications);

            // 按钮和版本信息面板
            Panel bottomPanel = new Panel();
            bottomPanel.Location = new Point(25, 170);
            bottomPanel.Size = new Size(550, 80);
            bottomPanel.BackColor = Color.Transparent;
            grpSettings.Controls.Add(bottomPanel);

            // 保存设置按钮
            btnSaveSettings = CreateStyledButton("保存设置", 0, 0, 140, 45);
            btnSaveSettings.Font = new Font("微软雅黑", 10, FontStyle.Bold);
            btnSaveSettings.Click += SaveSettings_Click;
            bottomPanel.Controls.Add(btnSaveSettings);

            // 版本信息
            lblVersion = new Label();
            lblVersion.Text = "自动开关机管理器 v3.1\n© 2024 All Rights Reserved";
            lblVersion.Location = new Point(200, 10);
            lblVersion.Size = new Size(350, 60);
            lblVersion.ForeColor = mutedTextColor;
            lblVersion.Font = new Font("微软雅黑", 9);
            lblVersion.TextAlign = ContentAlignment.MiddleRight;
            bottomPanel.Controls.Add(lblVersion);

            // 设置说明
            Label lblSettingsInfo = new Label();
            lblSettingsInfo.Text = "开机自启动功能将在用户登录时自动启动本程序，方便管理定时任务";
            lblSettingsInfo.Location = new Point(25, 310);
            lblSettingsInfo.Size = new Size(550, 30);
            lblSettingsInfo.ForeColor = mutedTextColor;
            lblSettingsInfo.Font = new Font("微软雅黑", 8);
            lblSettingsInfo.TextAlign = ContentAlignment.MiddleCenter;
            tabSettings.Controls.Add(lblSettingsInfo);
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

            chkAutoStart.TabIndex = 0;
            chkStartMinimized.TabIndex = 1;
            chkShowNotifications.TabIndex = 2;
            btnSaveSettings.TabIndex = 3;
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
            trayIcon.Text = "自动开关机管理器 v3.1";
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
                    if (showNotifications)
                    {
                        trayIcon.ShowBalloonTip(1000, "自动开关机管理器", "程序已最小化到系统托盘", ToolTipIcon.Info);
                    }
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

        private void SaveSettings_Click(object sender, EventArgs e)
        {
            SaveSettings();
            ApplySettings();
            ShowNotification("设置已保存成功！");
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
                        taskService.RootFolder.DeleteTask(TASK_BOOT, false);
                    }
                    catch
                    {
                        // 任务不存在，忽略错误
                    }

                    if (enabled)
                    {
                        // 创建开机任务
                        TaskDefinition bootTask = taskService.NewTask();
                        bootTask.RegistrationInfo.Description = "自动开机任务 - 自动开关机管理器";
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

                        // 设置操作 - 使用rundll32调用电源管理功能
                        bootTask.Actions.Add(new ExecAction("rundll32.exe", "powrprof.dll,SetSuspendState 0", null));
                        bootTask.Settings.WakeToRun = true;
                        bootTask.Settings.DisallowStartIfOnBatteries = false;
                        bootTask.Settings.StopIfGoingOnBatteries = false;
                        bootTask.Settings.ExecutionTimeLimit = TimeSpan.Zero;

                        taskService.RootFolder.RegisterTaskDefinition(TASK_BOOT, bootTask);

                        string repeatText = repeat ? "每天" : "一次";
                        string message = $"定时开机已设置为 {bootTime:yyyy-MM-dd HH:mm} ({repeatText})";
                        ShowNotification(message);

                        // 刷新任务列表
                        LoadStageTasks();
                    }
                    else
                    {
                        ShowNotification("开机任务已禁用");
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                LogError($"设置开机任务失败: {ex.Message}");
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
                    taskService.RootFolder.DeleteTask(TASK_BOOT, false);
                    ShowNotification("开机任务已删除");

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
                LogError($"删除开机任务失败: {ex.Message}");
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
                    // 使用shutdown命令设置定时关机，添加警告提示
                    string arguments = $"/s /t {seconds} /c \"自动开关机管理器：系统将在 {timeUntilShutdown:mm} 分 {timeUntilShutdown:ss} 秒后关机，请保存您的工作。\"";

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "shutdown",
                        Arguments = arguments,
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });

                    string message = $"系统将在 {shutdownTime:yyyy-MM-dd HH:mm} 关机";
                    ShowNotification(message);
                }
                else
                {
                    CancelShutdownTask();
                }

                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                LogError($"设置关机任务失败: {ex.Message}");
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

                ShowNotification("关机任务已取消");
                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                LogError($"取消关机任务失败: {ex.Message}");
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
                        Arguments = "/s /t 0 /c \"自动开关机管理器：立即关机\"",
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });
                }
                catch (Exception ex)
                {
                    LogError($"关机失败: {ex.Message}");
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
                        Arguments = "/r /t 0 /c \"自动开关机管理器：立即重启\"",
                        UseShellExecute = true,
                        Verb = IsRunningAsAdministrator() ? "runas" : ""
                    });
                }
                catch (Exception ex)
                {
                    LogError($"重启失败: {ex.Message}");
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
                        taskService.RootFolder.DeleteTask(TASK_STAGE_BOOT, false);
                        taskService.RootFolder.DeleteTask(TASK_STAGE_SHUTDOWN, false);
                    }
                    catch
                    {
                        // 任务不存在，忽略错误
                    }

                    if (chkStageEnabled.Checked)
                    {
                        // 创建阶段开机任务
                        TaskDefinition stageBootTask = taskService.NewTask();
                        stageBootTask.RegistrationInfo.Description = "阶段开机任务 - 自动开关机管理器";
                        stageBootTask.Principal.RunLevel = TaskRunLevel.Highest;

                        // 为每个选中的星期创建触发器
                        foreach (int dayIndex in chkListDays.CheckedIndices)
                        {
                            stageBootTask.Triggers.Add(new WeeklyTrigger
                            {
                                DaysOfWeek = (DaysOfTheWeek)(1 << dayIndex),
                                StartBoundary = DateTime.Today + dtpStageBootTime.Value.TimeOfDay,
                                Enabled = true
                            });
                        }

                        // 设置操作
                        stageBootTask.Actions.Add(new ExecAction("rundll32.exe", "powrprof.dll,SetSuspendState 0", null));
                        stageBootTask.Settings.WakeToRun = true;
                        stageBootTask.Settings.DisallowStartIfOnBatteries = false;
                        stageBootTask.Settings.StopIfGoingOnBatteries = false;
                        stageBootTask.Settings.ExecutionTimeLimit = TimeSpan.Zero;

                        taskService.RootFolder.RegisterTaskDefinition(TASK_STAGE_BOOT, stageBootTask);

                        // 创建阶段关机任务
                        TaskDefinition stageShutdownTask = taskService.NewTask();
                        stageShutdownTask.RegistrationInfo.Description = "阶段关机任务 - 自动开关机管理器";
                        stageShutdownTask.Principal.RunLevel = TaskRunLevel.Highest;

                        // 为每个选中的星期创建触发器
                        foreach (int dayIndex in chkListDays.CheckedIndices)
                        {
                            stageShutdownTask.Triggers.Add(new WeeklyTrigger
                            {
                                DaysOfWeek = (DaysOfTheWeek)(1 << dayIndex),
                                StartBoundary = DateTime.Today + dtpStageShutdownTime.Value.TimeOfDay,
                                Enabled = true
                            });
                        }

                        // 设置关机操作
                        stageShutdownTask.Actions.Add(new ExecAction("shutdown", "/s /t 0", null));

                        taskService.RootFolder.RegisterTaskDefinition(TASK_STAGE_SHUTDOWN, stageShutdownTask);

                        // 获取选择的星期文本
                        string daysText = GetSelectedDaysText();

                        string message = $"阶段任务已设置:\n开机: {dtpStageBootTime.Value:HH:mm}\n关机: {dtpStageShutdownTime.Value:HH:mm}\n执行: {daysText}";
                        ShowNotification(message);

                        // 刷新任务列表
                        LoadStageTasks();
                    }
                    else
                    {
                        ShowNotification("阶段任务已禁用");
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowAdminRequiredMessage();
            }
            catch (Exception ex)
            {
                LogError($"设置阶段任务失败: {ex.Message}");
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
                    taskService.RootFolder.DeleteTask(TASK_STAGE_BOOT, false);
                    taskService.RootFolder.DeleteTask(TASK_STAGE_SHUTDOWN, false);
                    ShowNotification("阶段任务已删除");

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
                LogError($"删除阶段任务失败: {ex.Message}");
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
                    Task stageBootTask = taskService.FindTask(TASK_STAGE_BOOT);
                    if (stageBootTask != null)
                    {
                        ListViewItem item = new ListViewItem("阶段开机");
                        item.SubItems.Add(dtpStageBootTime.Value.ToString("HH:mm"));
                        item.SubItems.Add(GetSelectedDaysText());
                        item.SubItems.Add(stageBootTask.Enabled ? "已启用" : "已禁用");
                        listStageTasks.Items.Add(item);
                    }

                    // 检查阶段关机任务
                    Task stageShutdownTask = taskService.FindTask(TASK_STAGE_SHUTDOWN);
                    if (stageShutdownTask != null)
                    {
                        ListViewItem item = new ListViewItem("阶段关机");
                        item.SubItems.Add(dtpStageShutdownTime.Value.ToString("HH:mm"));
                        item.SubItems.Add(GetSelectedDaysText());
                        item.SubItems.Add(stageShutdownTask.Enabled ? "已启用" : "已禁用");
                        listStageTasks.Items.Add(item);
                    }

                    // 检查单次开机任务
                    Task bootTask = taskService.FindTask(TASK_BOOT);
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

                    // 如果没有任务，显示提示
                    if (listStageTasks.Items.Count == 0)
                    {
                        ListViewItem item = new ListViewItem("暂无任务");
                        item.SubItems.Add("");
                        item.SubItems.Add("");
                        item.SubItems.Add("");
                        listStageTasks.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"加载任务列表失败: {ex.Message}");
            }
        }

        // 设置管理方法
        private void LoadSettings()
        {
            try
            {
                // 从注册表加载设置
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\AutoPowerManager", false))
                {
                    if (key != null)
                    {
                        autoStart = Convert.ToBoolean(key.GetValue("AutoStart", false));
                        startMinimized = Convert.ToBoolean(key.GetValue("StartMinimized", false));
                        showNotifications = Convert.ToBoolean(key.GetValue("ShowNotifications", true));
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"加载设置失败: {ex.Message}");
            }
        }

        private void SaveSettings()
        {
            try
            {
                // 更新设置变量
                autoStart = chkAutoStart.Checked;
                startMinimized = chkStartMinimized.Checked;
                showNotifications = chkShowNotifications.Checked;

                // 保存到注册表
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\AutoPowerManager"))
                {
                    key.SetValue("AutoStart", autoStart);
                    key.SetValue("StartMinimized", startMinimized);
                    key.SetValue("ShowNotifications", showNotifications);
                }
            }
            catch (Exception ex)
            {
                LogError($"保存设置失败: {ex.Message}");
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySettings()
        {
            // 应用开机自启动设置
            UpdateAutoStart();

            // 如果设置了启动时最小化，则最小化到托盘
            if (startMinimized)
            {
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
        }

        private void UpdateAutoStart()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (chkAutoStart.Checked)
                    {
                        // 添加开机自启动
                        key.SetValue("AutoPowerManager", $"\"{Application.ExecutablePath}\" /minimized");
                    }
                    else
                    {
                        // 移除开机自启动
                        key.DeleteValue("AutoPowerManager", false);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"设置开机自启动失败: {ex.Message}");
                MessageBox.Show($"设置开机自启动失败: {ex.Message}\n\n请尝试以管理员权限运行程序。", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                chkAutoStart.Checked = false;
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
            return selectedDays.Count > 0 ? string.Join(", ", selectedDays) : "无";
        }

        private string GetShutdownStatus()
        {
            try
            {
                // 使用更可靠的方法检测关机状态
                Process process = new Process();
                process.StartInfo.FileName = "shutdown";
                process.StartInfo.Arguments = "/a";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                // 检查多种可能的输出
                if (output.Contains("注销被取消") || output.Contains("shutdown was cancelled") ||
                    process.ExitCode == 1116 || output.ToLower().Contains("cancelled"))
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

        private void ShowNotification(string message)
        {
            if (showNotifications)
            {
                trayIcon?.ShowBalloonTip(3000, "自动开关机管理器", message, ToolTipIcon.Info);
            }
        }

        private void LogError(string message)
        {
            try
            {
                string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AutoPowerManager");
                Directory.CreateDirectory(logPath);
                string logFile = Path.Combine(logPath, "error.log");
                File.AppendAllText(logFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
            }
            catch
            {
                // 忽略日志错误
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
                if (showNotifications)
                {
                    trayIcon.ShowBalloonTip(1000, "自动开关机管理器", "程序已最小化到系统托盘", ToolTipIcon.Info);
                }
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