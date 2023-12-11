using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ScreenRestriction.Common;

namespace ScreenRestriction.Business_Objects
{
    [FormAttribute("USRSCRN", "Business_Objects/FrmUserConfiguration.b1f")]
    class FrmUserConfiguration : UserFormBase
    {
        public FrmUserConfiguration()
        {
        }
        public static SAPbouiCOM.Form objform;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsDetails1;
        private string strSQL;
        private SAPbobsCOM.Recordset objRs;
        SAPbouiCOM.ISBOChooseFromListEventArg pCFL;
        SAPbouiCOM.Column column;
        SAPbouiCOM.ComboBox comboBox;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lcode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tuserid").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tcode").Specific));
            this.EditText1.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText1_KeyDownAfter);
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tname").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkcode").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fldrlist").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("fldr2").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxlist").Specific));
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.ValidateBefore += new SAPbouiCOM._IMatrixEvents_ValidateBeforeEventHandler(this.Matrix0_ValidateBefore);
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void OnCustomInitialize()
        {
            try
            {
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_USRSCRN");
                odbdsDetails1 = objform.DataSources.DBDataSources.Item("@AT_USRSCRN1");
                Folder0.Item.Click();
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tuserid", false, true, false);
                objform.EnableMenu("1283", false);
                odbdsHeader.SetValue("Code", 0, clsModule.objaddon.USERNAME);
                if (((SAPbouiCOM.EditText)objform.Items.Item("tcode").Specific).String == "")
                    ((SAPbouiCOM.EditText)objform.Items.Item("tcode").Specific).String = clsModule.objaddon.USERNAME; // objaddon.objcompany.UserName
                if (clsModule.objaddon.HANA == true)
                {
                    ((SAPbouiCOM.EditText)objform.Items.Item("tname").Specific).String = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"U_NAME\" from OUSR where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'");
                    odbdsHeader.SetValue("U_UserId", 0, clsModule.objaddon.objglobalmethods.getSingleValue("Select \"USERID\" from OUSR where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'"));
                }
                else
                {
                    ((SAPbouiCOM.EditText)objform.Items.Item("tname").Specific).String = clsModule.objaddon.objglobalmethods.getSingleValue("Select U_NAME from OUSR where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'");
                    odbdsHeader.SetValue("U_UserId", 0, clsModule.objaddon.objglobalmethods.getSingleValue("Select USERID from OUSR where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'"));
                }

                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                Matrix0.Columns.Item("ltscrn").Visible = false;
                Matrix0.Columns.Item("sapscrn").Visible = false;
                Matrix0.Columns.Item("fielddesc").Visible = false;
                LoadCombo();

                if (EditText1.Value != "")
                {
                    if (clsModule.objaddon.HANA == true)
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 \"Status\" from \"@AT_USRSCRN\" where \"Code\"='" + EditText1.Value + "' ");
                    else
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 Status from [@AT_USRSCRN] where Code='" + EditText1.Value + "' ");

                    if (strSQL == "1")
                    {
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        if (clsModule.objaddon.HANA == true)
                            EditText0.Value = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"USERID\" from OUSR where \"USER_CODE\"='" + EditText1.Value + "'");//clsModule.objaddon. USERNAME;
                        else
                            EditText0.Value = clsModule.objaddon.objglobalmethods.getSingleValue("Select USERID from OUSR where USER_CODE='" + EditText1.Value + "'");//clsModule.objaddon. USERNAME;
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return;
                    }
                }


            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #region Fields

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText3;

        #endregion

        #region Form Events

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("USRSCRN", pVal.FormTypeCount);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "sapscrn", "#"); //fieldname
                ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(Matrix0.VisualRowCount).Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Data_Load_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #endregion

        #region Header Events

        private void EditText1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false)
                    return;
                pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pCFL.SelectedObjects != null)
                {
                    try
                    {
                        odbdsHeader.SetValue("U_UserId", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("USERID").Cells.Item(0).Value));
                        odbdsHeader.SetValue("Code", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("USER_CODE").Cells.Item(0).Value));
                        odbdsHeader.SetValue("Name", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_NAME").Cells.Item(0).Value));

                    }
                    catch (Exception ex)
                    {
                    }

                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void EditText1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                if (pVal.CharPressed == 9 & EditText1.Value != "")
                {
                    if (clsModule.objaddon.HANA == true)
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 \"Status\" from \"@AT_USRSCRN\" where \"Code\"='" + EditText1.Value + "' ");
                    else
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 Status from [@AT_USRSCRN] where Code='" + EditText1.Value + "' ");

                    if (strSQL == "1")
                    {
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        EditText0.Value = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"USERID\" from OUSR where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'");//clsModule.objaddon. USERNAME;
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return;
                    }
                    LoadCombo();
                }

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("E_Key_Down_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.InnerEvent == true) return;
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                RemoveLastrow(Matrix0, "fieldname");
                //if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(Matrix0.VisualRowCount).Specific).Selected != null & ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(Matrix0.VisualRowCount).Specific).String == "")
                //{
                //    clsModule.objaddon.objapplication.StatusBar.SetText("Select the Field Name on Line: " + Matrix0.VisualRowCount, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    //Matrix0.Columns.Item("fieldname").Cells.Item(Matrix0.VisualRowCount).Click();
                //    BubbleEvent = false; return;
                //}
                if (EditText1.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("User Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; return;
                }
                //for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                //{
                //    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String!= "") continue ;
                //    if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(i).Specific).Selected != null & ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String == "")
                //    {                        
                //            clsModule.objaddon.objapplication.StatusBar.SetSystemMessage("Select the Field Name on Line: " + i , SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //        Matrix0.Columns.Item("fieldname").Cells.Item(i).Click();
                //            return;                        
                //    }
                //}

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Click_Before: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #endregion

        #region Matrix CFL Screens



        private void Matrix0_ValidateBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //try
            //{
            //    string FieldName = ""; string ListScreen = ""; string MainScreen = "";
            //    if (pVal.InnerEvent == true) return;
            //    switch (pVal.ColUID)
            //    {
            //        case "fieldname":                        
            //            int curRow;
            //            int PrevRow = pVal.Row - 1;
            //            if (PrevRow == 0) return;

            //            //if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(PrevRow).Specific).Selected != null & ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(PrevRow).Specific).String == "")
            //            //{
            //            //    clsModule.objaddon.objapplication.StatusBar.SetText("Select the Field Name on Line: " + PrevRow, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //            //    //Matrix0.SetCellFocus(PrevRow, 5);
            //            //    //Matrix0.Columns.Item("fieldname").Cells.Item(PrevRow).Click();
            //            //    BubbleEvent = false; if (BubbleEvent == false) return;
            //            //}
            //            curRow = pVal.Row;
            //            if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String != "")
            //            {
            //                ListScreen = ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(curRow).Specific).Selected.Value;
            //                MainScreen = ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(curRow).Specific).Selected.Value;
            //                FieldName = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String;
            //            }

            //            for (int i = Matrix0.VisualRowCount; i >= 1; i--)
            //            {
            //                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String == "") continue;
            //                strSQL = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String;
            //                if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(i).Specific).Selected != null & ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(i).Specific).Selected != null)
            //                {
            //                    if (curRow != i & ListScreen == ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(i).Specific).Selected.Value & MainScreen == ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(i).Specific).Selected.Value & FieldName == ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String)
            //                    {
            //                        clsModule.objaddon.objapplication.StatusBar.SetSystemMessage("Duplicate Value ("+ strSQL + ") Found on line " + i + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String= "";
            //                        BubbleEvent = false; return;
            //                    }
            //                }
            //            }
            //            clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "fieldname", "#");
            //            ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(Matrix0.VisualRowCount).Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            //            break;
            //    }

            //}
            //catch (Exception ex)
            //{
            //    clsModule.objaddon.objapplication.StatusBar.SetText("Validate" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //}
        }

        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.InnerEvent == true) return;
                switch (pVal.ColUID)
                {
                    case "fieldname":
                        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(pVal.Row).Specific).String == "") ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("active").Cells.Item(pVal.Row).Specific).Checked = false;
                        if (pVal.CharPressed == 9 | pVal.CharPressed == 8 | pVal.CharPressed == 36) return;
                        BubbleEvent = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "cfl":
                        comboBox = (SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(pVal.Row).Specific;
                        Matrix0.SetCellWithoutValidation(pVal.Row, "ltscrn", comboBox.Selected.Description);
                        break;
                    case "mainscrn":
                        comboBox = (SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(pVal.Row).Specific;
                        Matrix0.SetCellWithoutValidation(pVal.Row, "sapscrn", comboBox.Selected.Description);
                        break;
                }



            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Combo_Select_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "fieldname":
                        if (pVal.CharPressed == 9 & pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)
                        {
                            if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(pVal.Row).Specific).Selected is null)
                            {
                                Matrix0.Columns.Item("cfl").Cells.Item(pVal.Row).Click(); return;
                            }
                            else if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(pVal.Row).Specific).Selected is null)
                            {
                                Matrix0.Columns.Item("mainscrn").Cells.Item(pVal.Row).Click(); return;
                            }
                            if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(pVal.Row).Specific).String != "") return;
                            clsModule.objaddon.frmdataselectform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            clsModule.objaddon.dataselRow = pVal.Row;
                            //clsModule.objaddon.TABLENAME = Convert.ToString(((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(pVal.Row).Specific).Selected.Value);
                            FrmSingleSelect activeform = new FrmSingleSelect();
                            activeform.Show();
                            activeform.LoadData(Convert.ToString(((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(pVal.Row).Specific).Selected.Value), EditText1.Value, Convert.ToString(((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(pVal.Row).Specific).Selected.Value));
                        }
                        break;
                }


            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("M_Key_Down_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                string FieldName = ""; string ListScreen = ""; string MainScreen = "";
                if (pVal.InnerEvent == true) return;
                switch (pVal.ColUID)
                {
                    case "fieldname":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "fieldname", "#");
                        ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(Matrix0.VisualRowCount).Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        int curRow;
                        int PrevRow = pVal.Row - 1;
                        if (PrevRow == 0) return;
                        curRow = pVal.Row;
                        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String != "")
                        {
                            ListScreen = ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(curRow).Specific).Selected.Value;
                            MainScreen = ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(curRow).Specific).Selected.Value;
                            FieldName = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String;
                        }

                        for (int i = Matrix0.VisualRowCount; i >= 1; i--)
                        {
                            if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String == "") continue;
                            strSQL = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String;
                            if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(i).Specific).Selected != null & ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(i).Specific).Selected != null)
                            {
                                if (curRow != i & ListScreen == ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(i).Specific).Selected.Value & MainScreen == ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("mainscrn").Cells.Item(i).Specific).Selected.Value & FieldName == ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(i).Specific).String)
                                {
                                    clsModule.objaddon.objapplication.StatusBar.SetSystemMessage("Duplicate Value (" + strSQL + ") Found on line " + i + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("fieldname").Cells.Item(curRow).Specific).String = "";
                                    return;
                                }
                            }
                        }

                        break;
                }

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Validate" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Functions

        private void LoadCombo()
        {
            try
            {
                if (EditText0.Value == "")
                    return;
                //objaddon.objapplication.StatusBar.SetText("Loading Item Group & Warehouse Details. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_USRSCRN");
                odbdsDetails1 = objform.DataSources.DBDataSources.Item("@AT_USRSCRN1");
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objform.Freeze(true);
                if (clsModule.objaddon.HANA)
                {
                    strSQL = "Select 'OITM' \"Table Name\",'List of Items' \"Screen Title\" from Dummy";
                    objRs.DoQuery(strSQL);
                }
                else
                {
                    strSQL = "Select 'OITM' [Table Name],'List of Items' [Screen Title] ";
                    objRs.DoQuery(strSQL);
                }


                Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "fieldname", "#");
                column = Matrix0.Columns.Item("cfl");
                if (column.ValidValues.Count == 0)
                {
                    while (!objRs.EoF)
                    {
                        column.ValidValues.Add(Convert.ToString(objRs.Fields.Item("Table Name").Value), Convert.ToString(objRs.Fields.Item("Screen Title").Value));
                        objRs.MoveNext();
                    }
                }

                ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("cfl").Cells.Item(Matrix0.VisualRowCount).Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                if (clsModule.objaddon.HANA)
                {
                    strSQL = "Select * from (Select '4' \"Object Type\", 'Items' \"Screen Name\" from dummy Union All";
                    strSQL += "\n Select '13' , 'A/R Invoice' from dummy Union All Select '14' , 'A/R Credit Memo' from dummy Union All";
                    strSQL += "\n Select '15' , 'Delivery' from dummy Union All Select '16' , 'Return' from dummy Union All";
                    strSQL += "\n Select '17' , 'Sales Order' from dummy Union All Select '18' , 'A/P Invoice' from dummy Union All";
                    strSQL += "\n Select '19' , 'A/P Credit Memo' from dummy Union All Select '20' , 'Goods Receipt PO' from dummy Union All";
                    strSQL += "\n Select '21' , 'Goods Return' from dummy Union All Select '22' , 'Purchase Order' from dummy Union All";
                    strSQL += "\n Select '23' , 'Sales Quotation' from dummy Union All Select '59' , 'Goods Receipt' from dummy Union All";
                    strSQL += "\n Select '60' , 'Goods Issue' from dummy Union All Select '66' , 'Bill of Materials' from dummy Union All";
                    strSQL += "\n Select '67' , 'Inventory Transfer' from dummy Union All Select '112' , 'Documents - Drafts' from dummy Union All";
                    strSQL += "\n Select '162' , 'Inventory Revaluation' from dummy Union All Select '202' , 'Production Order' from dummy Union All";
                    strSQL += "\n Select '203' , 'A/R Down Payment' from dummy Union All Select '204' , 'A/P Down Payment' from dummy Union All";
                    strSQL += "\n Select '1250000001' , 'Inventory Transfer Request' from dummy Union All Select '234000031' , 'Return Request' from dummy Union All";
                    strSQL += "\n Select '234000032' , 'Goods Return Request' from dummy Union All Select '' , '' from dummy) A Order by Cast(A.\"Object Type\" as Integer)";

                    objRs.DoQuery(strSQL);
                }
                else
                {
                    strSQL = "Select * from (Select '4' [Object Type], 'Items' [Screen Name]  Union All";
                    strSQL += "\n Select '13' , 'A/R Invoice'  Union All Select '14' , 'A/R Credit Memo'  Union All";
                    strSQL += "\n Select '15' , 'Delivery'  Union All Select '16' , 'Return'  Union All";
                    strSQL += "\n Select '17' , 'Sales Order'  Union All Select '18' , 'A/P Invoice'  Union All";
                    strSQL += "\n Select '19' , 'A/P Credit Memo'  Union All Select '20' , 'Goods Receipt PO'  Union All";
                    strSQL += "\n Select '21' , 'Goods Return'  Union All Select '22' , 'Purchase Order'  Union All";
                    strSQL += "\n Select '23' , 'Sales Quotation'  Union All Select '59' , 'Goods Receipt'  Union All";
                    strSQL += "\n Select '60' , 'Goods Issue'  Union All Select '66' , 'Bill of Materials'  Union All";
                    strSQL += "\n Select '67' , 'Inventory Transfer'  Union All Select '112' , 'Documents - Drafts'  Union All";
                    strSQL += "\n Select '162' , 'Inventory Revaluation'  Union All Select '202' , 'Production Order'  Union All";
                    strSQL += "\n Select '203' , 'A/R Down Payment'  Union All Select '204' , 'A/P Down Payment'  Union All";
                    strSQL += "\n Select '1250000001' , 'Inventory Transfer Request'  Union All Select '234000031' , 'Return Request'  Union All";
                    strSQL += "\n Select '234000032' , 'Goods Return Request' Union All Select '' , '' ) A Order by Cast(A.[Object Type] as Integer)";

                    objRs.DoQuery(strSQL);
                }
                column = Matrix0.Columns.Item("mainscrn");
                if (column.ValidValues.Count == 0)
                {
                    while (!objRs.EoF)
                    {
                        column.ValidValues.Add(Convert.ToString(objRs.Fields.Item("Object Type").Value), Convert.ToString(objRs.Fields.Item("Screen Name").Value));
                        objRs.MoveNext();
                    }
                }
                Matrix0.AutoResizeColumns();
                //clsModule. objaddon.objapplication.StatusBar.SetText("Item Group & Warehouse Details loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                objform.Freeze(false);
                objRs = null;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText("LoadCombo: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
        {
            try
            {
                if (omatrix.VisualRowCount == 0)
                    return;
                if (string.IsNullOrEmpty(Columname_check.ToString()))
                    return;
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific).String == "")
                {
                    omatrix.DeleteRow(omatrix.VisualRowCount);
                }
            }
            catch (Exception ex)
            {

            }

            #endregion
        }

    }
}
