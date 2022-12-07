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
            this.lbCardCode = ((SAPbouiCOM.StaticText)(this.GetItem("lbCardCode").Specific));
            this.txtCardCode = ((SAPbouiCOM.EditText)(this.GetItem("tCardCode").Specific));
            this.txtCardCode.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tCardName").Specific));
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.btnSearch = ((SAPbouiCOM.Button)(this.GetItem("btnSearch").Specific));
            this.btnSearch.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.gridQuery = ((SAPbouiCOM.Grid)(this.GetItem("grdOPCH").Specific));
            this.gridManual = ((SAPbouiCOM.Grid)(this.GetItem("grdMOPCH").Specific));
            this.oMatrix = ((SAPbouiCOM.Matrix)(this.GetItem("Item_8").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.oOptionButtonN = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_12").Specific));
            this.oOptionButtonN.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn0_ClickBefore);
            this.oOptionButtonP = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_13").Specific));
            this.oOptionButtonP.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn1_ClickBefore);
            this.oOptionButtonDate = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_14").Specific));
            this.oOptionButtonDate.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.OptionBtn2_ClickBefore);
            this.oOptionButtonP.GroupWith("Item_12");
            this.oOptionButtonDate.GroupWith("Item_12");
            this.PictureLAB = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_15").Specific));
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

        private void OnCustomInitialize()
        {
            oOptionButtonN.Selected = true;
        }

        private SAPbouiCOM.StaticText lbCardCode;
        private SAPbouiCOM.EditText txtCardCode;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.Button btnSearch;
        private SAPbouiCOM.Grid gridQuery;
        private SAPbouiCOM.Grid gridManual;

        private StaticText StaticText2;
        private StaticText StaticText3;
        private StaticText StaticText4;
        private EditText EditText2;
        private EditText EditText3;
        private StaticText StaticText5;
        private StaticText StaticText6;
        private OptionBtn oOptionButtonN;
        private OptionBtn oOptionButtonP;
        private OptionBtn oOptionButtonDate;

        private PictureBox PictureLAB;
        private StaticText StaticText7;

        private SAPbouiCOM.DataTable dt_OCRD;
        private SAPbouiCOM.DataTable dt_Matrix;
        private SAPbouiCOM.DataTable dt_Status;


        private SAPbouiCOM.Matrix oMatrix;

        private Form oForm;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            dt_Matrix = UIAPIRawForm.DataSources.DataTables.Item("DS_Matrix");
            dt_Status = UIAPIRawForm.DataSources.DataTables.Item("DS_Status");

            string queryRelatorio = "SELECT top 10 * FROM V_FISCAL ";

            string queryPV = "SELECT top 10 DocEntry, CardCode, CardName, DocDate, DocDueDate, DocTotal FROM ORDR ";

            if (!string.IsNullOrEmpty(txtCardCode.Value))
            {
                queryRelatorio += (queryRelatorio.Contains("WHERE") ? "AND " : "WHERE ") + $"\"Codigo PN\" = '{txtCardCode.Value}'";
                queryPV += (queryPV.Contains("WHERE") ? "AND " : "WHERE ") + $"\"CardCode\" = '{txtCardCode.Value}'";
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

            gridQuery.DataTable.ExecuteQuery(queryRelatorio);

            Recordset oRecord = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecord.DoQuery(queryPV);

            int index = 0;

            UIAPIRawForm.Freeze(true);
            ClearDataTable(gridManual.DataTable);
            while (!oRecord.EoF)
            {
                gridManual.DataTable.Rows.Add();

                gridManual.DataTable.Columns.Item("Documento").Cells.Item(index).Value = oRecord.Fields.Item("DocEntry").Value.ToString();
                gridManual.DataTable.Columns.Item("Cliente Código").Cells.Item(index).Value = oRecord.Fields.Item("CardCode").Value.ToString();
                gridManual.DataTable.Columns.Item("Cliente Nome").Cells.Item(index).Value = oRecord.Fields.Item("CardName").Value.ToString();
                gridManual.DataTable.Columns.Item("Data Documento").Cells.Item(index).Value = oRecord.Fields.Item("DocDate").Value;
                gridManual.DataTable.Columns.Item("Data Vencimento").Cells.Item(index).Value = oRecord.Fields.Item("DocDueDate").Value;
                gridManual.DataTable.Columns.Item("Total").Cells.Item(index).Value = oRecord.Fields.Item("DocTotal").Value;

                oRecord.MoveNext();
                index++;
            }
            UIAPIRawForm.Freeze(false);

            gridQuery.AutoResizeColumns();
            gridManual.AutoResizeColumns();

            dt_Matrix.ExecuteQuery("SELECT top 100 CardCode, DocDate, CONCAT(Min(DocDate), ' - ', Max(DocDate)) AS 'Periodo',(SUM(OI.DocTotal) / MONTH((Min(DocDate) - Max(DocDate)))) AS 'FatMes',(SUM(OI.DocTotal) / YEAR((Min(DocDate) - Max(DocDate)))) AS 'FatAno',SUM(OI.DocTotal) AS 'Total' FROM OINV OI GROUP BY CardCode, DocDate");

            for (int i = 0; i < dt_Matrix.Rows.Count; i++)
            {
                dt_Status.Rows.Add();

                dt_Status.Columns.Item("Status").Cells.Item(i).Value = i % 2 == 0 ? "Atrasada" : "Correta";
            }

            oMatrix.LoadFromDataSource();

            for (int i = 0; i < oMatrix.RowCount; i++)
            {
                if (dt_Status.Columns.Item("Status").Cells.Item(i).Value.ToString() == "Atrasada")
                {
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 6, RGBtoInt(255, 160, 122));
                }
                else
                {
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 6, RGBtoInt(124, 252, 0));
                }
            }
        }

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

        private void OptionBtn0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            gridManual.CollapseLevel = 0;
        }

        private void OptionBtn1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            gridManual.CollapseLevel = 2;
        }

        private void OptionBtn2_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var userDS = oForm.DataSources.UserDataSources.Item("OpBtnDS");

            string b = userDS.Value;

            gridManual.CollapseLevel = 3;
        }

        public static int RGBtoInt(int r, int g, int b)
        {
            return (r << 0) | (g << 8) | (b << 16);
        }


    }
}