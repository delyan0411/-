using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Aspose.Cells;
using System.Reflection;
using MySql.Data.MySqlClient;
using System.Data;

namespace 账单明细导出.Controllers
{
    /// <summary>
    /// 
    /// </summary>
    public class Zdyssj
    {
        public string 保险公司 { get; set; }
        public string 企业 { get; set; }
        public string 门店 { get; set; }
        public string 姓名 { get; set; }
        public string 身份证 { get; set; }
        public string 流水号 { get; set; }
        public decimal 交易金额 { get; set; }
        public string 交易日期 { get; set; }
        public string 交易类型 { get; set; }
        public string 交易状态 { get; set; }
        //保险公司 企业  门店 姓名  身份证 流水号 交易金额 交易日期    交易类型 交易状态
    }
    /// <summary>
    /// 
    /// </summary>
    public class Splb
    {
        public string product_name { get; set; }
        public decimal sale_price { get; set; }
        public uint product_type_id { get; set; }
        public string product_type_path { get; set; }
        //product_name,sale_price,product_type_id,product_type_path
    }
    /// <summary>
    /// 
    /// </summary>
    public class Xsmx
    {
        public string order_no { get; set; }
        public DateTime pay_time { get; set; }
        public int product_id { get; set; }
        public string product_name { get; set; }
        public decimal deal_price { get; set; }
        public int sale_num { get; set; }
        public uint product_type_id { get; set; }
        public string product_type_path { get; set; }
        public decimal zje { get; set; }
        //order_no,pay_time, product_id, product_name, deal_price, sale_num, product_type_id, product_type_path

    }
    /// <summary>
    /// 
    /// </summary>
    public class Ypmxqd
    {
        public string 企业 { get; set; }
        public string 门店名称 { get; set; }
        public string 姓名 { get; set; }
        public string 身份证号 { get; set; }
        public string 交易日期 { get; set; }
        public decimal 交易金额 { get; set; }
        public string 交易状态 { get; set; }
        public string 开票日期 { get; set; }
        public string 发票号码 { get; set; }
        public string 药品清单 { get; set; }
        public decimal 单价 { get; set; }
        public int 数量 { get; set; }
        public decimal 金额 { get; set; }
        public string 日期 { get; set; }
        //企业  门店名称 姓名  身份证号 交易日期    交易金额 交易状态    开票日期 发票号码    药品清单 单价  数量 金额  日期
    }
    /// <summary>
    /// 
    /// </summary>
    public class Check
    {
        public string 订单号 { get; set; }
        public string 是否存在 { get; set; }
    }

    public class Clzd
    {
        int[] canpay = { 1, 2 };
        List<Splb> splb = new List<Splb>();
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public List<T> DtConvertToList<T>(DataTable dt) where T : new()
        {
            // 定义集合  
            List<T> ts = new List<T>();
            // 获得此模型的类型  
            Type type = typeof(T);
            //定义一个临时变量  
            string tempName = string.Empty;
            //遍历DataTable中所有的数据行  
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                // 获得此模型的公共属性  
                PropertyInfo[] propertys = t.GetType().GetProperties();
                //遍历该对象的所有属性  
                foreach (PropertyInfo pi in propertys)
                {
                    tempName = pi.Name;//将属性名称赋值给临时变量  
                    //检查DataTable是否包含此列（列名==对象的属性名）    
                    if (dt.Columns.Contains(tempName))
                    {
                        // 判断此属性是否有Setter  
                        if (!pi.CanWrite) continue;//该属性不可写，直接跳出  
                        //取值  
                        object value = dr[tempName];
                        //如果非空，则赋给对象的属性  
                        if (value != DBNull.Value)
                            pi.SetValue(t, value, null);
                    }
                }
                //对象添加到泛型集合中  
                ts.Add(t);
            }
            return ts;
        }

        /// <summary>
        /// 从Excel获取数据
        /// </summary>
        /// <typeparam name="T">对象</typeparam>
        /// <param name="filePath">文件完整路径</param>
        /// <returns>对象列表</returns>
        public List<T> GetObjectList<T>(Worksheet sheet) where T : new()
        {
            List<T> list = new List<T>();
            //if (!filePath.Trim().EndsWith("csv") && !filePath.Trim().EndsWith("xlsx"))
            //{
            //    return list;
            //}

            Type type = typeof(T);
            //Workbook workbook = new Workbook(filePath);
            //Worksheet sheet = workbook.Worksheets[0];
            // 获取标题列表    
            var titleDic = this.GetTitleDic(sheet);
            // 循环每行数据    
            for (int i = 1; i < int.MaxValue; i++)
            {// 行为空时结束
                if (string.IsNullOrEmpty(sheet.Cells[i, 0].StringValue))
                { break; }
                T instance = new T();
                // 循环赋值每个属性
                foreach (var item in type.GetProperties())
                {
                    if (titleDic.ContainsKey(item.Name))
                    {
                        string str = sheet.Cells[i, titleDic[item.Name]].StringValue;
                        if (!string.IsNullOrEmpty(str))
                        {
                            try
                            {
                                // 根据类型进行转换赋值
                                if (item.PropertyType == typeof(string))
                                {
                                    item.SetValue(instance, str);
                                }
                                else if (item.PropertyType.IsEnum)
                                {
                                    item.SetValue(instance, int.Parse(str));
                                }
                                else
                                {
                                    MethodInfo method = item.PropertyType.GetMethod("Parse", new Type[] { typeof(string) });
                                    object obj = null;
                                    if (method != null)
                                    {
                                        obj = method.Invoke(null, new object[] { str });
                                        item.SetValue(instance, obj);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                // 获取错误  
                            }
                        }
                    }
                }
                list.Add(instance);
            }
            return list;
        }

        /// <summary>
        /// 把对象List保存到Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objList"></param>
        /// <param name="saveFilePath"></param>
        public Worksheet SetExcelList<T>(List<T> objList, Worksheet sheet)
        {
            // 冻结第一行    
            sheet.FreezePanes(1, 1, 1, 0);
            // 循环插入每行    
            int row = 0;
            foreach (var obj in objList)
            {
                int column = 0;
                var properties = obj.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.DeclaredOnly);
                if (row == 0)
                {
                    foreach (var titName in properties)
                    {
                        sheet.Cells[0, column].PutValue(titName.Name);
                        sheet.Cells.SetRowHeight(0, 30);
                        Style style = sheet.Cells[0, column].GetStyle();
                        style.ForegroundColor = System.Drawing.Color.FromArgb(128, 128, 128);
                        style.Pattern = BackgroundType.Solid;
                        style.Font.Color = System.Drawing.Color.White;
                        sheet.Cells[0, column].SetStyle(style);
                        column++;

                    }
                    row++;
                }
                // 循环插入当前行的每列
                column = 0;
                foreach (var property in properties)
                {
                    var itemValue = property.GetValue(obj);
                    if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(int))
                    {
                        sheet.Cells[row, column].PutValue(itemValue.ToString(), true);
                        Style style = sheet.Cells[row, column].GetStyle();
                        //设置cell大小 设置背景颜色 最后增加一行总结
                        //保单号 药品明细清单
                        style.Number = 0;
                        sheet.Cells[row, column].SetStyle(style);
                        if (itemValue.ToString() == "0")
                        {
                            sheet.Cells[row, column].PutValue("");
                        }
                    }
                    else
                    {
                        sheet.Cells[row, column].PutValue(itemValue.ToString());
                    }
                    column++;
                }
                row++;
            }
            return sheet;
        }

        /// <summary>
        /// 获取标题行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Dictionary<string, int> GetTitleDic(Worksheet sheet)
        {
            Dictionary<string, int> titList = new Dictionary<string, int>();
            for (int i = 0; i < int.MaxValue; i++)
            {
                if (sheet.Cells[0, i].StringValue == string.Empty) { return titList; }
                titList.Add(sheet.Cells[0, i].StringValue, i);
            }
            return titList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lzdlist"></param>
        /// <returns></returns>
        public Workbook clexcel(List<Zdyssj> lzdlist, string starttime, string endtime)
        {
            Workbook workbookret = new Workbook();
            WorksheetCollection worksheets = workbookret.Worksheets;
            worksheets.Clear();
            //连接mysql 然后获取商品列表Server=myServerAddress;Database=myDataBase;Uid=myUsername;Pwd=myPassword;
            string str = "Server=120.55.186.172;User ID=root;Password=jiuzhou!2016@#;Database=jz_shop;";
            MySqlConnection con = new MySqlConnection(str);
            //实例化链接
            con.Open();
            //开启连接
            string strpro = "select product_name,sale_price,product_type_id,product_type_path from sp_product where is_on_sale=1 and is_visible=1  and total_sale_count>0";
            MySqlCommand cmdpro = new MySqlCommand(strpro, con);
            MySqlDataAdapter adapro = new MySqlDataAdapter(cmdpro);
            DataSet dspro = new DataSet();
            adapro.Fill(dspro);
            splb = DtConvertToList<Splb>(dspro.Tables[0]);

            List<Splb> delsplb = new List<Splb>();
            foreach (var spitem in splb)
            {
                int mainpath = 0;
                try
                {
                    mainpath = Convert.ToInt16(spitem.product_type_path.Split(',')[1]);
                }
                catch
                {
                    mainpath = 0;
                }
                if (Array.IndexOf(canpay, mainpath) == -1)
                {
                    delsplb.Add(spitem);
                }
            }
            foreach (var spitem in delsplb)
            {
                splb.Remove(spitem);
            }
            //查询结果填充数据集
            int pay_type = 53;
            //string starttime = "2017-05-01";
            //string endtime = "2017-07-01";
            //vod.pay_type = {2} and 
            string strorder = string.Format("SELECT vod.order_no,vod.pay_time,vod.product_id,vod.product_name,vod.deal_price,vod.sale_num,vod.order_money+vod.trans_money as zje,sp.product_type_id,sp.product_type_path FROM view_order_detail as vod INNER JOIN sp_product as sp on sp.product_id = vod.product_id where  vod.pay_state = 2 and vod.pay_time > '{0}' and vod.pay_time < '{1}' ORDER BY vod.pay_time,vod.order_no", starttime, endtime);
            MySqlCommand cmdorder = new MySqlCommand(strorder, con);
            MySqlDataAdapter adaorder = new MySqlDataAdapter(cmdorder);
            DataSet dsorder = new DataSet();
            adaorder.Fill(dsorder);
            List<Xsmx> xsmx = DtConvertToList<Xsmx>(dsorder.Tables[0]);
            //查询结果填充数据集
            con.Close();
            //关闭连接
            //遍历导入的excel 每一万元的list生成sheet ,清空list
            decimal totalmoney = 0;
            List<Ypmxqd> ypmxqdlist = new List<Ypmxqd>();
            List<Check> checklist = new List<Check>();
            int sheeti = 10000;
            int counti = 0;
            foreach (var item in lzdlist)
            {
                bool existxs = xsmx.Any(t => t.order_no == item.流水号.Replace("jz", ""));
                //查看有没有这个订单 没有跳过(???) 

                if (item.交易金额 >= 10000)
                {
                    sheeti = sheeti + 1;
                    Worksheet worksheetret = worksheets.Add("发票号000" + sheeti + "");
                    worksheetret.AutoFitColumns();
                    //生成sheet ,清空list
                    //不符合 相近的价格的 top 然后价格和数量取原来的
                    if (ypmxqdlist.Count > 0)
                    {
                        SetExcelList(ypmxqdlist, worksheetret);
                    }
                    totalmoney = 0;
                    ypmxqdlist.Clear();
                    if (existxs)
                    {
                        additems(xsmx, ypmxqdlist, item);
                    }
                    sheeti = sheeti + 1;
                    worksheetret = worksheets.Add("发票号000" + sheeti + "");
                    SetExcelList(ypmxqdlist, worksheetret);
                    ypmxqdlist.Clear();
                }
                else
                {
                    decimal isover = totalmoney + item.交易金额;
                    if (isover > 10000)
                    {
                        sheeti = sheeti + 1;
                        Worksheet worksheetret = worksheets.Add("发票号000" + sheeti + "");
                        worksheetret.AutoFitColumns();
                        //生成sheet ,清空list
                        SetExcelList(ypmxqdlist, worksheetret);
                        totalmoney = 0;
                        ypmxqdlist.Clear();
                    }
                    totalmoney = totalmoney + item.交易金额;
                    if (existxs)
                    {
                        additems(xsmx, ypmxqdlist, item);
                    }
                }
                counti = counti + 1;
                if (counti == lzdlist.Count)
                {
                    sheeti = sheeti + 1;
                    Worksheet worksheetret = worksheets.Add("发票号000" + sheeti + "");
                    SetExcelList(ypmxqdlist, worksheetret);
                }

                Check c = new Check();
                c.订单号 = item.流水号;
                c.是否存在 = existxs ? "是" : "否否否";
                checklist.Add(c);
            }
            Worksheet worksheetcheck = worksheets.Add("检查excel");
            SetExcelList(checklist, worksheetcheck);
            //写入输出book 
            return workbookret;
        }



       /// <summary>
       /// 
       /// </summary>
       /// <param name="lzdlist"></param>
       /// <param name="starttime"></param>
       /// <param name="endtime"></param>
       /// <returns></returns>
        public Workbook clexcelall(List<Zdyssj> lzdlist, string starttime, string endtime)
        {
            Workbook workbookret = new Workbook();
            WorksheetCollection worksheets = workbookret.Worksheets;
            worksheets.Clear();
            //连接mysql 然后获取商品列表Server=myServerAddress;Database=myDataBase;Uid=myUsername;Pwd=myPassword;
            string str = "Server=120.55.186.172;User ID=root;Password=jiuzhou!2016@#;Database=jz_shop;";
            MySqlConnection con = new MySqlConnection(str);
            //实例化链接
            con.Open();
            //开启连接
            string strpro = "select product_name,sale_price,product_type_id,product_type_path from sp_product where is_on_sale=1 and is_visible=1  and total_sale_count>0";
            MySqlCommand cmdpro = new MySqlCommand(strpro, con);
            MySqlDataAdapter adapro = new MySqlDataAdapter(cmdpro);
            DataSet dspro = new DataSet();
            adapro.Fill(dspro);
            splb = DtConvertToList<Splb>(dspro.Tables[0]);

            List<Splb> delsplb = new List<Splb>();
            foreach (var spitem in splb)
            {
                int mainpath = 0;
                try
                {
                    mainpath = Convert.ToInt16(spitem.product_type_path.Split(',')[1]);
                }
                catch
                {
                    mainpath = 0;
                }
                if (Array.IndexOf(canpay, mainpath) == -1)
                {
                    delsplb.Add(spitem);
                }
            }
            foreach (var spitem in delsplb)
            {
                splb.Remove(spitem);
            }
            //查询结果填充数据集
            int pay_type = 53;
            //string starttime = "2017-05-01";
            //string endtime = "2017-07-01";
            //vod.pay_type = {2} and 
            string strorder = string.Format("SELECT vod.order_no,vod.pay_time,vod.product_id,vod.product_name,vod.deal_price,vod.sale_num,vod.order_money+vod.trans_money as zje,sp.product_type_id,sp.product_type_path FROM view_order_detail as vod INNER JOIN sp_product as sp on sp.product_id = vod.product_id where  vod.pay_state = 2 and vod.pay_time > '{0}' and vod.pay_time < '{1}' ORDER BY vod.pay_time,vod.order_no", starttime, endtime);
            MySqlCommand cmdorder = new MySqlCommand(strorder, con);
            MySqlDataAdapter adaorder = new MySqlDataAdapter(cmdorder);
            DataSet dsorder = new DataSet();
            adaorder.Fill(dsorder);
            List<Xsmx> xsmx = DtConvertToList<Xsmx>(dsorder.Tables[0]);
            //查询结果填充数据集
            con.Close();
            //关闭连接
            //遍历导入的excel 每一万元的list生成sheet ,清空list
            List<Ypmxqd> ypmxqdlist = new List<Ypmxqd>();
            List<Check> checklist = new List<Check>();
            foreach (var item in lzdlist)
            {
                bool existxs = xsmx.Any(t => t.order_no == item.流水号.Replace("jz", ""));
                //查看有没有这个订单 没有跳过(???)             
                if (existxs)
                {
                    additems(xsmx, ypmxqdlist, item);
                }
                Check c = new Check();
                c.订单号 = item.流水号;
                c.是否存在 = existxs ? "是" : "否否否";
                checklist.Add(c);
            }
            Worksheet worksheetall = worksheets.Add("订单列表");
            SetExcelList(ypmxqdlist, worksheetall);
            Worksheet worksheetcheck = worksheets.Add("检查excel");
            SetExcelList(checklist, worksheetcheck);
            //写入输出book 
            return workbookret;
        }

        /// <summary>
        /// 
        /// </summary>
        public void additems(List<Xsmx> xsmx, List<Ypmxqd> ypmxqdlist, Zdyssj item)
        {

            List<Xsmx> oneorderxsmx = xsmx.Where(t => t.order_no == item.流水号.Replace("jz", "")).ToList();
            int mxi = 0;
            decimal lastje = 0;
            //查询明细  
            foreach (var mxitem in oneorderxsmx)
            {
                mxi = mxi + 1;
                Ypmxqd ypmxqd = new Ypmxqd();
                if (mxi == 1)
                {
                    ypmxqd.企业 = item.企业;
                    ypmxqd.门店名称 = item.门店;
                    ypmxqd.姓名 = item.姓名;
                    ypmxqd.身份证号 = item.身份证;
                    ypmxqd.交易日期 = item.交易日期;
                    ypmxqd.交易金额 = item.交易金额;
                    ypmxqd.交易状态 = "交易完成";
                    ypmxqd.开票日期 = "";
                    ypmxqd.发票号码 = "";
                }
                else
                {
                    ypmxqd.企业 = "";
                    ypmxqd.门店名称 = "";
                    ypmxqd.姓名 = "";
                    ypmxqd.身份证号 = "";
                    ypmxqd.交易日期 = "";
                    ypmxqd.交易金额 = 0;
                    ypmxqd.交易状态 = "";
                    ypmxqd.开票日期 = "";
                    ypmxqd.发票号码 = "";
                }
                int mainpath = 0;
                try
                {
                    mainpath = Convert.ToInt16(mxitem.product_type_path.Split(',')[1]);
                }
                catch
                {
                    mainpath = 0;
                }
                //不符合 相近的价格的 top 然后价格和数量取原来的
                if (Array.IndexOf(canpay, mainpath) == -1)
                {
                    //替换不符合的明细  
                    Splb sp = splb.OrderBy(t => Math.Abs(t.sale_price - mxitem.deal_price)).First();
                    ypmxqd.药品清单 = sp.product_name;
                }
                //符合购买限制
                else
                {
                    ypmxqd.药品清单 = mxitem.product_name;
                }
                ypmxqd.单价 = mxitem.deal_price;
                ypmxqd.数量 = mxitem.sale_num;

                if (mxi == oneorderxsmx.Count)
                {
                    ypmxqd.金额 = item.交易金额 - lastje;
                }
                else
                {
                    ypmxqd.金额 = mxitem.deal_price * mxitem.sale_num;
                }
                lastje += (mxitem.deal_price * mxitem.sale_num);

                ypmxqd.日期 = item.交易日期;
                //插入到list一条数据
                ypmxqdlist.Add(ypmxqd);
            }

        }
    }
    public class HomeController : Controller
    {
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase fexcel, string date_s, string date_e,int op)
        {
            ViewData["tips"] = "";
            if (fexcel == null)
            {
                ViewData["tips"] = "必须上传一个excel文件";
            }
            else
            {
                string[] headerarr = { "保险公司", "企业", "门店", "姓名", "身份证", "流水号", "交易金额", "交易日期", "交易类型", "交易状态" };
                //var fileName = fexcel.FileName;
                var fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(fexcel.FileName);
                var exportexlname = "明细(" + fexcel.FileName + ")";
                var filePath = Server.MapPath(string.Format("~/{0}", "File"));
                if (Path.GetExtension(fileName) == ".xls" || Path.GetExtension(fileName) == ".xlsx")
                {
                    fexcel.SaveAs(Path.Combine(filePath, fileName));
                    //获取到excel

                    Zdyssj zd = new Zdyssj();
                    Workbook workbook = new Workbook(Path.Combine(filePath, fileName));
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;
                    //判断excel格式
                    for (int i = 0; i < 10; i++)
                    {
                        if (cells[0, i].StringValue.Trim() != headerarr[i])
                        {
                            ViewData["tips"] = "excel字段不对";
                            return View(ViewData["tips"]);
                        }
                    }
                    //去除空行 并写入list
                    Clzd clzd = new Clzd();
                    List<Zdyssj> lzdlist = clzd.GetObjectList<Zdyssj>(worksheet);
                    //Instantiating a Workbook object
                    string starttime = string.IsNullOrEmpty(date_s) ? DateTime.Now.AddMonths(-3).ToString("yyyy-MM-01") : date_s;
                    string endtime = string.IsNullOrEmpty(date_e) ? DateTime.Now.AddMonths(-1).ToString("yyyy-MM-01") : date_e;
                    Workbook wk = new Workbook();
                    if (op == 1)
                    {
                        wk = clzd.clexcel(lzdlist, starttime, endtime);
                    }
                    else
                    {
                       wk = clzd.clexcelall(lzdlist, starttime, endtime);
                    }
                    //输出这个excel
                    return File(wk.SaveToStream().ToArray(), "application/ms-excel", "" + exportexlname + ".xls");
                }
                else
                {
                    ViewData["tips"] = "必须上传一个excel文件";
                }
            }
            return View(ViewData["tips"]);


        }

        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

    }
}