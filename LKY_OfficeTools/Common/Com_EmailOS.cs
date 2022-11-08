/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_EmailOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Documents;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 发送邮件的类库
    /// </summary>
    internal class Com_EmailOS
    {
        /*
        /// <summary>
        /// 通过本机无密码发送邮件
        /// </summary>
        /// <returns></returns>
        internal static bool Send_Localhost()
        {
            MailMessage msg = new MailMessage();

            //收件人邮箱
            msg.To.Add("liukaiyuan@sjtu.edu.cn");

            //抄送人邮箱
            //msg.CC.Add("c@c.com");

            //发件人信息：邮箱地址（可以随便写）、发件人名字、编码
            msg.From = new MailAddress("a@a.com", "AlphaWu", Encoding.UTF8);

            //邮件标题 & 编码
            msg.Subject = "这是测试邮件";
            msg.SubjectEncoding = Encoding.UTF8;

            //邮件内容 & 编码
            msg.Body = "邮件内容";
            msg.BodyEncoding = Encoding.UTF8;

            //是否是HTML邮件
            msg.IsBodyHtml = false;

            //邮件优先级 
            msg.Priority = MailPriority.High;

            //开始发送
            SmtpClient client = new SmtpClient();
            client.Host = "localhost";              //使用本机发送

            try
            {
                //object userState = msg;
                //client.SendAsync(msg, userState);   //异步

                client.Send(msg);

                return true;
            }
            catch (SmtpException ex)
            {
                new Log(ex);
                return false;
            }
        }
        */

        /// <summary>
        /// 通过账号方式发送
        /// </summary>
        /// <param name="send_from_mail">发送者邮箱</param>
        /// <param name="send_from_username">发送人名称</param>
        /// <param name="send_to_mail">收件人邮箱</param>
        /// <param name="mail_subject">主题</param>
        /// <param name="mail_body">内容</param>
        /// <param name="mail_file">附件</param>
        /// <param name="SMTPHost">smtp服务器</param>
        /// <param name="SMTPuser">邮箱</param>
        /// <param name="SMTPpass">密码</param>
        /// <param name="priority">邮件优先级，默认为一般，设置为高时，收件人看到标题旁边，会有红色叹号</param>
        /// <returns></returns>
        public static bool Send_Account( string send_from_username, string send_to_mail, string mail_subject,
            string mail_body, List<string> mail_file, string SMTPHost, string SMTPuser, string SMTPpass, MailPriority priority = MailPriority.Normal)
        {
            //设置from和to地址
            MailAddress from = new MailAddress(SMTPuser, send_from_username);
            MailAddress to = new MailAddress(send_to_mail);

            //创建一个MailMessage对象
            MailMessage oMail = new MailMessage(from, to);

            try
            {
                // 添加附件
                if (mail_file != null)
                {
                    foreach (var now_file in mail_file)
                    {
                        if (File.Exists(now_file))
                        {
                            oMail.Attachments.Add(new Attachment(now_file));
                        }
                    }
                }

                //邮件标题
                oMail.Subject = mail_subject;

                //设置邮件为html格式
                oMail.IsBodyHtml = true;

                //邮件内容
                oMail.Body = mail_body;

                //邮件格式
                oMail.IsBodyHtml = true;

                //邮件采用的编码
                oMail.BodyEncoding = Encoding.UTF8;

                //设置邮件的优先级
                oMail.Priority = priority;

                //发送邮件
                SmtpClient client = new SmtpClient();
                //client.UseDefaultCredentials = false;
                client.Host = SMTPHost;

                //企业邮箱需设置端口，个人邮箱不需要
                client.EnableSsl = true;    //开启SSL加密
                client.Port = 587;
                client.Credentials = new NetworkCredential(SMTPuser, SMTPpass);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                client.Send(oMail);

                return true;
            }
            catch /*(Exception err)*/
            {
                //new Log(err);
                //Console.ReadKey();
                return false;
            }
            finally
            {
                //释放资源
                oMail.Dispose();
            }
        }
    }
}
