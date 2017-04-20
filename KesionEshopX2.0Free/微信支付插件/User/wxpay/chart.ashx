<%@ WebHandler Language="C#" Class="chart" %>

using System;
using System.Web;
using ThoughtWorks.QRCode.Codec;

public class chart : IHttpHandler {

    public void ProcessRequest (HttpContext context) {
         if (context.Request["text"] == null) { return; }
            QRCodeEncoder qrCodeEncoder = new QRCodeEncoder();
            try
            {
                int scale = Convert.ToInt16(context.Request["size"].ToString());
                qrCodeEncoder.QRCodeScale = scale;
            }
            catch { }
            String data = context.Request["text"].ToString();
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            System.Drawing.Image myimg = qrCodeEncoder.Encode(data, System.Text.Encoding.UTF8); //kedee 增加utf-8编码，可支持中文汉字  
            myimg.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
       
            context.Response.ClearContent();
            context.Response.ContentType = "image/Gif";
            context.Response.BinaryWrite(ms.ToArray());
            context.Response.End();
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}