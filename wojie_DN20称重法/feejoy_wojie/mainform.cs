using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using feejoy_wojie.database;
using DevExpress.LookAndFeel;
using WHC.Framework.BaseUI;
using WHC.Framework.Commons;

namespace feejoy_wojie
{
    public partial class mainform : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public mainform()
        {
            InitializeComponent();
            UserLookAndFeel.Default.SetSkinStyle("Office 2019 Colorful");
        }
        private void mainform_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;

            //获取屏幕的宽度和高度
            int w = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
            int h = System.Windows.Forms.SystemInformation.VirtualScreen.Height;

            //设置最大尺寸  和  最小尺寸  （如果没有修改默认值，则不用设置）
            this.MaximumSize = new Size(w, h);
            this.MinimumSize = new Size(w, h);

            //设置窗口位置
            this.Location = new Point(0, 0);

            //设置窗口大小
            this.Width = w;
            this.Height = h;

            ChildWinManagement.LoadMdiForm(this, typeof(feejoy_wojie.subform.monitor));
            this.xtraTabbedMdiManager1.ClosePageButtonShowMode = DevExpress.XtraTab.ClosePageButtonShowMode.InAllTabPagesAndTabControlHeader;
        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //eejoy_wojie.subform.monitor monitor = new subform.monitor();
            //monitor.ShowDialog();
            ChildWinManagement.LoadMdiForm(this, typeof(feejoy_wojie.subform.monitor));
            this.xtraTabbedMdiManager1.ClosePageButtonShowMode = DevExpress.XtraTab.ClosePageButtonShowMode.InAllTabPagesAndTabControlHeader;
        }

        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            subform.calibratedata calibratedata = new subform.calibratedata();
            calibratedata.ShowDialog();
        }

    }
}
