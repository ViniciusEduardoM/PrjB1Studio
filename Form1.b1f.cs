using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Xml;
using Application = SAPbouiCOM.Framework.Application;

namespace SAPExemple
{
    [FormAttribute("SAPExemple.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lbCardCode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tCardCode").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tCardName").Specific));
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btnSearch").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grdOPCH").Specific));
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grdMOPCH").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_8").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_12").Specific));
            this.OptionBtn0.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn0_ClickBefore);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_13").Specific));
            this.OptionBtn1.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn1_ClickBefore);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_14").Specific));
            this.OptionBtn2.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn2_ClickBefore);
            this.OptionBtn1.GroupWith("Item_12");
            this.OptionBtn2.GroupWith("Item_12");
            this.PictureBox1 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_15").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            OptionBtn0.Selected = true;
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;

        private SAPbouiCOM.DataTable dt_OCRD;


        private Form oForm;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            string queryRelatorio = "SELECT top 10 * FROM V_FISCAL ";

            string queryPV = "SELECT top 10 DocEntry, CardCode, CardName, DocDate, DocDueDate, DocTotal FROM ORDR ";

            if (!string.IsNullOrEmpty(EditText0.Value))
            {
                queryRelatorio += (queryRelatorio.Contains("WHERE") ? "AND " : "WHERE ") + $"\"Codigo PN\" = '{EditText0.Value}'";
                queryPV += (queryPV.Contains("WHERE") ? "AND " : "WHERE ") + $"\"CardCode\" = '{EditText0.Value}'";
            }

            if (!string.IsNullOrEmpty(EditText2.Value))
            {
                queryRelatorio += (queryRelatorio.Contains("WHERE") ? "AND " : "WHERE ") + $"\"Dt.Entrada\" >= '{EditText2.Value}'";
                queryPV += (queryPV.Contains("WHERE") ? "AND " : "WHERE ") + $"\"DocDate\" >= '{EditText2.Value}'";
            }

            if (!string.IsNullOrEmpty(EditText3.Value))
            {
                queryRelatorio += (queryRelatorio.Contains("WHERE") ? "AND " : "WHERE ") + $"\"Dt.Entrada\" <= '{EditText3.Value}'";
                queryPV += (queryPV.Contains("WHERE") ? "AND " : "WHERE ") + $"\"DocDate\" <= '{EditText3.Value}'";
            }

            Grid0.DataTable.ExecuteQuery(queryRelatorio);

            Recordset oRecord = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecord.DoQuery(queryPV);

            int index = 0;

            UIAPIRawForm.Freeze(true);
            ClearDataTable(Grid1.DataTable);
            while (!oRecord.EoF)
            {
                Grid1.DataTable.Rows.Add();

                Grid1.DataTable.Columns.Item("Documento").Cells.Item(index).Value = oRecord.Fields.Item("DocEntry").Value.ToString();
                Grid1.DataTable.Columns.Item("Cliente Código").Cells.Item(index).Value = oRecord.Fields.Item("CardCode").Value.ToString();
                Grid1.DataTable.Columns.Item("Cliente Nome").Cells.Item(index).Value = oRecord.Fields.Item("CardName").Value.ToString();
                Grid1.DataTable.Columns.Item("Data Documento").Cells.Item(index).Value = oRecord.Fields.Item("DocDate").Value;
                Grid1.DataTable.Columns.Item("Data Vencimento").Cells.Item(index).Value = oRecord.Fields.Item("DocDueDate").Value;
                Grid1.DataTable.Columns.Item("Total").Cells.Item(index).Value = oRecord.Fields.Item("DocTotal").Value;

                oRecord.MoveNext();
                index++;
            }
            UIAPIRawForm.Freeze(false);

            Grid0.AutoResizeColumns();
            Grid1.AutoResizeColumns();
        }

        private SAPbouiCOM.Matrix Matrix0;

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            ISBOChooseFromListEventArg oCFLEvento = ((ISBOChooseFromListEventArg)(pVal));

            SAPbouiCOM.DataTable dt = oCFLEvento.SelectedObjects;
            dt_OCRD = oForm.DataSources.DataTables.Item("DT_OCRD");

            dt_OCRD.ExecuteQuery($"SELECT * FROM OCRD WHERE CardCode = '{dt.GetValue(0, 0)}'");

            oForm.Update();

        }

        private void EditText1_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            EditText0_ChooseFromListAfter(sboObject, pVal);
        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            oForm = Application.SBO_Application.Forms.Item("Form1");
        }

        private void ClearDataTable(SAPbouiCOM.DataTable datatable)
        {
            var rowsCount = datatable.Rows.Count;
            for (int i = 0; i < rowsCount; i++)
            {
                datatable.Rows.Remove(0);
            }
        }

        private StaticText StaticText2;
        private StaticText StaticText3;
        private StaticText StaticText4;
        private EditText EditText2;
        private EditText EditText3;
        private StaticText StaticText5;
        private StaticText StaticText6;
        private OptionBtn OptionBtn0;
        private OptionBtn OptionBtn1;
        private OptionBtn OptionBtn2;

        private void OptionBtn0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            Grid1.CollapseLevel = 0;
        }

        private void OptionBtn1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            Grid1.CollapseLevel = 2;
        }

        private void OptionBtn2_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            Grid1.CollapseLevel = 3;
        }

        private PictureBox PictureBox1;
        private StaticText StaticText7;
    }
}