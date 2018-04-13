//
// Converter.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2015-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.IO;
using System.Text;
using MimeKit;
using MsgKit.Helpers;
using System.Linq;
using MsgKit.Exceptions;

namespace MsgKit
{
    /// <summary>
    ///     This class exposes some methods to convert MSG to EML and vice versa
    /// </summary>
    public static class Converter
    {
        #region ConvertEmlToMsg
        /// <summary>
        ///     Converts an EML file to MSG format
        /// </summary>
        /// <param name="emlFileName">The EML (MIME) file</param>
        private static Email ConvertEmlToMsgLib(string emlFileName)
        {
            var eml = MimeMessage.Load(emlFileName);
            var sender = new Sender(string.Empty, string.Empty);
            var FromAddress = ((MailboxAddress)eml.From[0]);

            if (eml.From.Count > 0)
                sender = new Sender(FromAddress.Address, FromAddress.Name);

            var representing = new Representing(FromAddress.Address, FromAddress.Name); //Default to from address
            if (eml.ResentSender != null)
                representing = new Representing(eml.ResentSender.Address, eml.ResentSender.Name);

            var msg = new Email(sender, representing, eml.Subject)
            {
                SentOn = eml.Date.UtcDateTime,
                InternetMessageId = eml.MessageId
            };

            switch (eml.Priority)
            {
                case MessagePriority.NonUrgent:
                    msg.Priority = Enums.MessagePriority.PRIO_NONURGENT;
                    break;
                case MessagePriority.Normal:
                    msg.Priority = Enums.MessagePriority.PRIO_NORMAL;
                    break;
                case MessagePriority.Urgent:
                    msg.Priority = Enums.MessagePriority.PRIO_URGENT;
                    break;
            }

            switch (eml.Importance)
            {
                case MessageImportance.Low:
                    msg.Importance = Enums.MessageImportance.IMPORTANCE_LOW;
                    break;
                case MessageImportance.Normal:
                    msg.Importance = Enums.MessageImportance.IMPORTANCE_NORMAL;
                    break;
                case MessageImportance.High:
                    msg.Importance = Enums.MessageImportance.IMPORTANCE_HIGH;
                    break;
            }

            foreach (var to in eml.To)
            {
                var mailAddress = (MailboxAddress)to;
                msg.Recipients.AddTo(mailAddress.Address, mailAddress.Name);
            }

            foreach (var cc in eml.Cc)
            {
                var mailAddress = (MailboxAddress)cc;
                msg.Recipients.AddCc(mailAddress.Address, mailAddress.Name);
            }

            foreach (var bcc in eml.Bcc)
            {
                var mailAddress = (MailboxAddress)bcc;
                msg.Recipients.AddBcc(mailAddress.Address, mailAddress.Name);
            }

            using (var headerStream = new MemoryStream())
            {
                eml.Headers.WriteTo(headerStream);
                headerStream.Position = 0;
                msg.TransportMessageHeaders = Encoding.ASCII.GetString(headerStream.ToArray());
            }

            msg.BodyHtml = eml.HtmlBody;
            msg.BodyText = eml.TextBody;

            try
            {
                var NamelessCount = 0;
                foreach (var bodyPart in eml.BodyParts)
                {
                    var type = bodyPart.GetType();
                    if (type.Equals(typeof(TextPart))) //Make sure the TextPart isnt TextBody or HtmlBody
                    {
                        var part = (TextPart)bodyPart;
                        if (new[] { eml.TextBody, eml.HtmlBody }.Contains(part.Text))
                        {
                            continue; //Break conversion
                        }
                        else if (part.ContentDisposition?.Disposition == ContentDisposition.Inline &&
                            string.IsNullOrEmpty(bodyPart.ContentId)) //If TextPart is inline but doesnt have content defined
                        {
                            msg.BodyText += Environment.NewLine + part.Text; //Append text to the body of the email
                            msg.BodyHtml += part.Text.Replace(Environment.NewLine, "<br>");
                            continue;
                        }
                    }

                    //Get the body part and convert as needed
                    var attachmentStream = new MemoryStream();
                    var fileName = bodyPart.ContentType.Name;
                    var extension = string.Empty;

                    if (type.Equals(typeof(MessagePart)))
                    {
                        var part = (MessagePart)bodyPart;
                        part.Message.WriteTo(attachmentStream);
                        if (part.Message != null)
                            fileName = part.Message.Subject;
                        extension = ".eml";
                    }
                    else if (type.Equals(typeof(MessageDispositionNotification)))
                    {
                        var part = (MessageDispositionNotification)bodyPart;
                        fileName = part.FileName;
                    }
                    else if (type.Equals(typeof(MessageDeliveryStatus)))
                    {
                        var part = (MessageDeliveryStatus)bodyPart;
                        fileName = "details";
                        extension = ".txt";
                        part.WriteTo(FormatOptions.Default, attachmentStream, true);
                    }
                    else
                    {
                        var part = (MimePart)bodyPart;
                        part.Content.DecodeTo(attachmentStream);
                        fileName = part.FileName;
                        bodyPart.WriteTo(attachmentStream);
                    }

                    fileName = string.IsNullOrWhiteSpace(fileName)
                        ? "Nameless" + (NamelessCount++ > 0 ? NamelessCount.ToString() : "")
                        : FileManager.RemoveInvalidFileNameChars(fileName);

                    if (!string.IsNullOrEmpty(extension))
                        fileName += extension;

                    var inline = bodyPart.ContentDisposition?.Disposition.Equals(
                        ContentDisposition.Inline,
                        StringComparison.InvariantCultureIgnoreCase) ?? false;

                    if (inline && string.IsNullOrEmpty(bodyPart.ContentId))
                        inline = false;

                    attachmentStream.Position = 0;
                    msg.Attachments.Add(attachmentStream, fileName, -1, inline, bodyPart.ContentId);
                }
            }
            catch
            {
                throw;
            }

            return msg;
        }

        /// <summary>
        ///     Converts an EML file to MSG format
        /// </summary>
        /// <param name="emlFileName">The EML (MIME) file</param>
        /// <param name="msgFileName">The MSG file</param>
        public static void ConvertEmlToMsg(string emlFileName, string msgFileName) =>
            ConvertEmlToMsgLib(emlFileName).Save(msgFileName);

        /// <summary>
        ///     Converts an EML file to MSG format
        /// </summary>
        /// <param name="emlFileName">The EML (MIME) file</param>
        public static Email ConvertEmlToMsg(string emlFileName) => 
            ConvertEmlToMsgLib(emlFileName);
        #endregion

        #region ConvertMsgToEml
        /// <summary>
        ///     Converts an MSG file to EML format
        /// </summary>
        /// <param name="msgFileName">The MSG file</param>
        /// <param name="emlFileName">The EML (MIME) file</param>
        public static void ConvertMsgToEml(string msgFileName, string emlFileName)
        {
            //var eml = MimeKit.MimeMessage.CreateFromMailMessage()
            throw new NotImplementedException("Not yet done");
        }
        #endregion
    }
}
