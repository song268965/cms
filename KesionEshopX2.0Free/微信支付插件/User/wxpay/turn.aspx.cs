using System;
using System.Web.UI;
using System.Xml;
using System.IO;
using WeiPay;


    public partial class WxPay : System.Web.UI.Page
    {
        //页面输出 不用操作
        public static string Code = "";     //微信端传来的code
        public static string PrepayId = ""; //预支付ID
        public static string Sign = "";     //为了获取预支付ID的签名
        public static string PaySign = "";  //进行支付需要的签名
        public static string Package = "";  //进行支付需要的包
        public static string TimeStamp = ""; //时间戳 程序生成 无需填写
        public static string NonceStr = ""; //随机字符串  程序生成 无需填写



        protected void Page_Load(object sender, EventArgs e)
        {


        #region 基本参数===========================


        //接收并读取POST过来的XML文件流
        StreamReader reader = new StreamReader(Request.InputStream);
        String xmlData = reader.ReadToEnd();
        //把数据重新返回给客户端
      //  Response.Write(xmlData);
       // Response.End();



            //Utils.WriteLog("WeiPay 页面  package（XML）：" + data);

            string prepayXml = HttpUtil.Send(xmlData, "https://api.mch.weixin.qq.com/pay/unifiedorder");
        // Utils.WriteLog("WeiPay 页面  package（Back_XML）：" + prepayXml);

      //  Response.Write(prepayXml);
      //  Response.End();

            //获取预支付ID
            var xdoc = new XmlDocument();
            xdoc.LoadXml(prepayXml);
            XmlNode xn = xdoc.SelectSingleNode("xml");
            XmlNodeList xnl = xn.ChildNodes;
            if (xnl.Count > 7)
            {
                PrepayId = xnl[7].InnerText;
                Package = string.Format("prepay_id={0}", PrepayId);
                // Utils.WriteLog("WeiPay 页面  package：" + Package);
            }
            #endregion
            Response.Write(PrepayId);
            Response.End();

        }



    }

