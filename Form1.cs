using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace Upgrade
{
    public partial class Form1 : Form
    {
        string 状态 = "准备中";
        string 下载链接 = "";
        string 目标文件 = "";
        int 完成倒计时 = 5;
        public Form1()
        {
            InitializeComponent();
            _ = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            await CheckUpgradeFileAsync();//加载升级配置文件
            await CheckAndKillTargetProcessAsync();//检查目标文件进程是否在运行，如果是则终止它
            await DownloadAndExtractFileAsync();//下载并解压文件
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (状态 == "完成")
            {
                try
                {
                    // 异步启动目标文件
                    if (!string.IsNullOrEmpty(目标文件))
                    {
                        状态 = $"正在启动{目标文件}...";

                        // 检查目标文件是否存在
                        string targetFilePath = System.IO.Path.Combine(Application.StartupPath, 目标文件);
                        if (System.IO.File.Exists(targetFilePath))
                        {
                            // 使用Process.Start异步启动应用程序
                            System.Diagnostics.Process.Start(targetFilePath);
                            状态 = $"已启动{目标文件}";
                        }
                        else
                        {
                            状态 = $"目标文件不存在: {目标文件}";
                            MessageBox.Show($"目标文件不存在: {targetFilePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    状态 = $"启动目标文件失败: {ex.Message}";
                    MessageBox.Show($"启动目标文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 无论启动是否成功，都关闭当前窗口
                    this.Close();
                }
            }
            else
            {
                this.Close();
            }   
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            groupBox1.Text = "状态：" + 状态;
        }

        private async Task CheckUpgradeFileAsync()
        {
            // 获取应用程序的当前目录
            string currentDirectory = Application.StartupPath;

            // 检查Upgrade.ini是否存在
            string upgradeFilePath = System.IO.Path.Combine(currentDirectory, "Upgrade.ini");

            if (System.IO.File.Exists(upgradeFilePath))
            {
                状态 = "升级配置文件已就绪";

                // 读取INI文件内容
                try
                {
                    // 使用异步方法读取文件
                    string[] lines;
                    using (var reader = new System.IO.StreamReader(upgradeFilePath))
                    {
                        string content = await reader.ReadToEndAsync();
                        lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                    }

                    string currentSection = "";

                    foreach (string line in lines)
                    {
                        string trimmedLine = line.Trim();

                        // 跳过空行和注释行
                        if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";"))
                            continue;

                        // 检查是否是节名
                        if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                        {
                            currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);
                            continue;
                        }

                        // 处理键值对
                        if (currentSection == "Upgrade" && trimmedLine.Contains("="))
                        {
                            int equalPos = trimmedLine.IndexOf('=');
                            string key = trimmedLine.Substring(0, equalPos).Trim();
                            string value = trimmedLine.Substring(equalPos + 1).Trim();

                            // 根据键名赋值给对应变量
                            if (key == "DownloadURL")
                            {
                                下载链接 = value;
                            }
                            else if (key == "TargetFile")
                            {
                                目标文件 = value;
                            }
                        }
                    }

                    // 检查是否成功获取到必要信息
                    if (string.IsNullOrEmpty(下载链接) || string.IsNullOrEmpty(目标文件))
                    {
                        状态 = "升级配置文件内容不完整";
                        MessageBox.Show("升级配置文件内容不完整，请检查。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    状态 = "读取升级配置文件失败";
                    MessageBox.Show("读取升级配置文件失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
            }
            else
            {
                状态 = "升级配置文件不存在";
                // 提示升级文件不存在
                MessageBox.Show("升级文件不存在，程序将退出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // 关闭窗体
                this.Close();
            }
        }

        /// <summary>
        /// 异步检查目标文件进程是否在运行，如果是则终止它
        /// </summary>
        private async Task CheckAndKillTargetProcessAsync()
        {
            try
            {
                // 确保目标文件名称不为空
                if (string.IsNullOrEmpty(目标文件))
                {
                    状态 = "目标文件名称未指定";
                    return;
                }

                状态 = "正在检查目标进程...";

                // 分离文件名（无路径）和文件扩展名
                string fileName = System.IO.Path.GetFileNameWithoutExtension(目标文件);

                // 使用Task.Run在后台线程执行可能耗时的进程操作
                await Task.Run(() =>
                {
                    try
                    {
                        // 获取与目标文件相同名称的所有进程
                        System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(fileName);

                        if (processes.Length > 0)
                        {
                            // 更新状态
                            状态 = $"发现{processes.Length}个{fileName}进程，正在终止...";

                            // 遍历找到的所有进程
                            foreach (System.Diagnostics.Process process in processes)
                            {
                                try
                                {
                                    // 尝试终止进程
                                    process.Kill();
                                    // 等待进程退出，最多等待5秒
                                    process.WaitForExit(5000);
                                }
                                catch (Exception ex)
                                {
                                    // 处理无法终止进程的异常
                                    状态 = $"终止进程失败: {ex.Message}";
                                    throw;  // 重新抛出异常以便外层捕获
                                }
                                finally
                                {
                                    // 无论成功与否，都要清理资源
                                    process.Dispose();
                                }
                            }

                            状态 = $"已成功终止{fileName}进程";
                        }
                        else
                        {
                            状态 = $"{fileName}进程未在运行";
                        }
                    }
                    catch (Exception ex)
                    {
                        状态 = $"检查进程时出错: {ex.Message}";
                        throw;  // 重新抛出异常以便外层捕获
                    }
                });
            }
            catch (Exception ex)
            {
                状态 = $"操作失败: {ex.Message}";
                MessageBox.Show($"检查或终止目标进程时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        /// <summary>
        /// 异步下载并解压文件，显示下载进度
        /// </summary>
        private async Task DownloadAndExtractFileAsync()
        {
            状态 = "开始下载";
            string zipFilePath = System.IO.Path.Combine(Application.StartupPath, "Upgrade.zip");

            try
            {
                // 检查并删除已存在的Upgrade.zip
                if (System.IO.File.Exists(zipFilePath))
                {
                    状态 = "正在删除旧的升级包";
                    System.IO.File.Delete(zipFilePath);
                }

                状态 = "准备下载文件";

                // 创建不使用系统代理的WebClient
                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    // 绕过系统代理设置
                    client.Proxy = null;

                    // 设置下载进度更新事件
                    client.DownloadProgressChanged += (sender, e) =>
                    {
                        // 在UI线程上更新进度条
                        this.Invoke(new Action(() =>
                        {
                            progressBar1.Value = e.ProgressPercentage;
                            状态 = $"下载中: {e.ProgressPercentage}% ({e.BytesReceived / 1024} KB / {e.TotalBytesToReceive / 1024} KB)";
                        }));
                    };

                    // 下载完成事件
                    client.DownloadFileCompleted += (sender, e) =>
                    {
                        if (e.Error != null)
                        {
                            状态 = $"下载错误: {e.Error.Message}";
                        }
                        else if (!e.Cancelled)
                        {
                            状态 = "下载完成";
                        }
                    };

                    状态 = "正在下载文件...";

                    // 异步下载文件
                    await client.DownloadFileTaskAsync(new Uri(下载链接), zipFilePath);

                    状态 = "下载完成，准备解压";

                    // 解压文件
                    await Task.Run(() =>
                    {
                        try
                        {
                            // 解压到当前目录，覆盖已有文件
                            状态 = "正在解压文件...";

                            // 使用ZipArchive解压文件
                            using (System.IO.Compression.ZipArchive archive = System.IO.Compression.ZipFile.OpenRead(zipFilePath))
                            {
                                foreach (System.IO.Compression.ZipArchiveEntry entry in archive.Entries)
                                {
                                    string destinationPath = System.IO.Path.Combine(Application.StartupPath, entry.FullName);
                                    string destinationDirectory = System.IO.Path.GetDirectoryName(destinationPath);

                                    // 确保目标目录存在
                                    if (!string.IsNullOrEmpty(destinationDirectory) && !System.IO.Directory.Exists(destinationDirectory))
                                    {
                                        System.IO.Directory.CreateDirectory(destinationDirectory);
                                    }

                                    // 如果是目录条目（名称以/结尾），则跳过
                                    if (entry.FullName.EndsWith("/") || entry.FullName.EndsWith("\\"))
                                        continue;

                                    // 提取文件，覆盖已存在的文件
                                    try
                                    {
                                        // 如果文件存在，先删除
                                        if (System.IO.File.Exists(destinationPath))
                                        {
                                            System.IO.File.Delete(destinationPath);
                                        }

                                        // 解压文件
                                        entry.ExtractToFile(destinationPath);
                                    }
                                    catch (System.IO.IOException)
                                    {
                                        // 可能由于文件被锁定等原因无法删除，尝试使用不同的方法
                                        using (var entryStream = entry.Open())
                                        using (var fileStream = new System.IO.FileStream(destinationPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                                        {
                                            entryStream.CopyTo(fileStream);
                                        }
                                    }
                                }
                            }

                            状态 = "解压完成";
                        }
                        catch (Exception ex)
                        {
                            状态 = $"解压错误: {ex.Message}";
                            throw;
                        }
                        finally
                        {
                            try
                            {
                                // 删除下载的zip文件
                                if (System.IO.File.Exists(zipFilePath))
                                {
                                    System.IO.File.Delete(zipFilePath);
                                    状态 = "升级包已清理，升级完成";
                                }
                            }
                            catch (Exception ex)
                            {
                                状态 = $"删除升级包失败: {ex.Message}";
                            }
                        }
                    });

                }
            }
            catch (Exception ex)
            {
                状态 = $"升级过程出错: {ex.Message}";
                MessageBox.Show($"升级过程中出现错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            timer2.Enabled = true;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            状态 = "完成";
            button1.Text = $"{状态}({完成倒计时})";
            完成倒计时--;
            if (完成倒计时 < 0) {
                timer2.Stop();
                timer2.Enabled = false;
                button1_Click(sender, e);
            }
        }
    }
}
