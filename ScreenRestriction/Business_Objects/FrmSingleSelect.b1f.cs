using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ScreenRestriction.Common;

namespace ScreenRestriction.Business_Objects
{
    [FormAttribute("DATASEL", "Business_Objects/FrmSingleSelect.b1f")]
    class FrmSingleSelect : UserFormBase
    {
        public static SAPbouiCOM.Form objform;
        private SAPbouiCOM.Matrix objmatrix;
        string strSQL;

        public FrmSingleSelect()
        {
            // trigger after OnCustomInitialize Event
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btnsel").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("gridsel").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lfind").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tfind").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {

        }

        #region Fields

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;

        #endregion


        private void OnCustomInitialize()
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("DATASEL", 1);
                objform.Title = "Choose From List";
                objform.Left = clsModule.objaddon.frmdataselectform.Left + 100;
                objform.Top = clsModule.objaddon.frmdataselectform.Top + 100;
                //if (clsModule.objaddon.TABLENAME!="") LoadData(clsModule.objaddon.TABLENAME, USER, OBJTYPE);
                clsModule.objaddon.bModal = true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        public void LoadData(string TableName, string UserName, string ObjType)
        {
            try
            {
                //SAPbobsCOM.SBObob objBridge;
                SAPbobsCOM.Recordset objRs;
                //objBridge = (SAPbobsCOM.SBObob)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //objRs = objBridge.GetTableFieldList(TableName);//"OITM"

                if (clsModule.objaddon.HANA == true)
                {
                    //strSQL = "Select COLUMN_NAME \"Field Name\" from TABLE_COLUMNS Where SCHEMA_NAME='" + clsModule.objaddon.objcompany.CompanyDB + "' and TABLE_NAME='" + TableName + "' Order by COLUMN_NAME"; //POSITION
                    strSQL = "Select COLUMN_NAME \"Field Name\" from TABLE_COLUMNS Where SCHEMA_NAME='" + clsModule.objaddon.objcompany.CompanyDB + "' and TABLE_NAME='" + TableName + "' ";
                    strSQL += "\n and COLUMN_NAME Not in (Select \"U_FieldName\" from \"@AT_USRSCRN1\" T0 join \"@AT_USRSCRN\" T1 On T0.\"Code\"=T1.\"Code\" WHERE T1.\"Code\"='" + UserName + "' and T0.\"U_TableName\"='" + TableName + "' and T0.\"U_ObjType\"='" + ObjType + "')";
                    strSQL += "\n Order by COLUMN_NAME";
                }                    
                else
                {
                    //strSQL = "Select COLUMN_NAME [Field Name] from INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME='" + TableName + "' Order by COLUMN_NAME";//ORDINAL_POSITION
                    strSQL = "Select COLUMN_NAME [Field Name] from INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME='" + TableName + "'";
                    strSQL += "\n and COLUMN_NAME not in (Select U_FieldName from [@AT_USRSCRN1] T0 join [@AT_USRSCRN] T1 On T0.Code=T1.Code WHERE T1.Code='" + UserName + "' and T0.U_TableName='" + TableName + "' and T0.U_ObjType='" + ObjType + "')";
                    strSQL += "\n Order by COLUMN_NAME";
                }
                    

                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_0");
                objform.DataSources.DataTables.Item("DT_0").ExecuteQuery(strSQL);
                Grid0.RowHeaders.TitleObject.Caption = "#";
                for (int i = 0; i < Grid0.Columns.Count; i++)
                {
                    Grid0.Columns.Item(i).TitleObject.Sortable = true;
                    Grid0.Columns.Item(i).Editable = false;
                }
                Grid0.Rows.SelectedRows.Add(0);
                Grid0.AutoResizeColumns();
                objRs = null;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Clear();
                Grid0.Rows.SelectedRows.Add(pVal.Row);
               
            }
            catch (Exception ex)
            {

            }

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {                
               // int vv = Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
                if (Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))!=-1)
                {
                    objmatrix =(SAPbouiCOM.Matrix)clsModule.objaddon.frmdataselectform.Items.Item("mtxlist").Specific;
                    ((SAPbouiCOM.EditText)objmatrix.Columns.Item("fieldname").Cells.Item(clsModule.objaddon.dataselRow).Specific).String = Convert.ToString(Grid0.DataTable.GetValue(0, Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))));
                }
                ((SAPbouiCOM.CheckBox)objmatrix.Columns.Item("active").Cells.Item(clsModule.objaddon.dataselRow).Specific).Checked = true;
                clsModule.objaddon.frmdataselectform = null;
                clsModule.objaddon.dataselRow = 0;
                objform.Close();


            }
            catch (Exception ex)
            {

            }

        }       

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                string FindString = ((SAPbouiCOM.EditText)objform.Items.Item("tfind").Specific).String.ToUpper();
                string FieldName;               

                int rowindex = Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
                if (pVal.CharPressed == 38 & pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//UP
                {
                    if (rowindex != 0) Grid0.Rows.SelectedRows.Add(rowindex - 1);
                }
                else if (pVal.CharPressed == 40 & pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//DOWN
                {
                    Grid0.Rows.SelectedRows.Add(rowindex + 1);
                }
                else
                {
                    if (FindString == "") return;
                    for (int i = 0; i < Grid0.DataTable.Rows.Count - 1; i++)
                    {
                        FieldName = Grid0.DataTable.GetValue(0, Grid0.GetDataTableRowIndex(i)).ToString().ToUpper();
                        if (FieldName.StartsWith(FindString))//| FieldName.Contains(FindString) | FieldName.EndsWith(FindString)
                        {
                            Grid0.Rows.SelectedRows.Add(i);
                            break;
                        }
                    }
                }


            }
            catch (Exception ex)
            {

            }

        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.Row != -1) return;
                if (pVal.Row == -1)
                {
                    Grid0.Item.Click();
                    Grid0.Rows.SelectedRows.Add(0);
                }
                else
                {
                    if (Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)) != -1)
                    {
                        objmatrix = (SAPbouiCOM.Matrix)clsModule.objaddon.frmdataselectform.Items.Item("mtxlist").Specific;
                        ((SAPbouiCOM.EditText)objmatrix.Columns.Item("fieldname").Cells.Item(clsModule.objaddon.dataselRow).Specific).String = Convert.ToString(Grid0.DataTable.GetValue(0, Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))));
                    }
                    ((SAPbouiCOM.CheckBox)objmatrix.Columns.Item("active").Cells.Item(clsModule.objaddon.dataselRow).Specific).Checked = true;
                    clsModule.objaddon.frmdataselectform = null;
                    clsModule.objaddon.dataselRow = 0;
                    objform.Close();
                }
               
            }
            catch (Exception ex)
            {

            }

        }

        
    }
}
