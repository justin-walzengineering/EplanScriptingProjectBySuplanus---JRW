/* ================================================================================================
 * ExtendedSettings
 * ================================================================================================
 * ibKastl GmbH
 * Dr.-Seitz-Str. 13a
 * 82418 Murnau
 * 08841 6286833
 * info@ibkastl.de
 * http://www.ibkastl.de
 * ================================================================================================
 * Version
 * 2.1
 * ================================================================================================
*/

using System;
using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;

namespace ibKastl.Scripts.Global.ExtendedSettings
{
    public class ExtendedSettings : Form
    {
        private Button btnCancel;
        private Button btnOk;
        private CheckBox chbCopySettingsPath;
        private ToolTip tt;
        private CheckBox chbMenuId;
        public Eplan.EplApi.Base.Settings settings = new Eplan.EplApi.Base.Settings();
        private CheckBox chbInplaceEditing;
        private CheckBox chbDontChangeSourceText;
        private CheckBox chbRemoveSelection;
		private CheckBox chbDebugScripts;

		#region Vom Windows Form-Designer generierter Code
		// ReSharper disable All
		/// <summary>
		/// Erforderliche Designervariable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen
        /// gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor
        /// geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOk = new System.Windows.Forms.Button();
			this.chbCopySettingsPath = new System.Windows.Forms.CheckBox();
			this.chbMenuId = new System.Windows.Forms.CheckBox();
			this.tt = new System.Windows.Forms.ToolTip(this.components);
			this.chbInplaceEditing = new System.Windows.Forms.CheckBox();
			this.chbDontChangeSourceText = new System.Windows.Forms.CheckBox();
			this.chbRemoveSelection = new System.Windows.Forms.CheckBox();
			this.chbDebugScripts = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(137, 161);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(120, 23);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "Abort";
			this.btnCancel.UseVisualStyleBackColor = true;
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnOk
			// 
			this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOk.Location = new System.Drawing.Point(11, 161);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(120, 23);
			this.btnOk.TabIndex = 0;
			this.btnOk.Text = "OK";
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// chbCopySettingsPath
			// 
			this.chbCopySettingsPath.AutoSize = true;
			this.chbCopySettingsPath.Location = new System.Drawing.Point(12, 12);
			this.chbCopySettingsPath.Name = "chbCopySettingsPath";
			this.chbCopySettingsPath.Size = new System.Drawing.Size(113, 17);
			this.chbCopySettingsPath.TabIndex = 2;
			this.chbCopySettingsPath.Text = "Copy settings path";
			this.tt.SetToolTip(this.chbCopySettingsPath, "Adds menu item in the context menu in the settings to copy the path" +
        "");
			this.chbCopySettingsPath.UseVisualStyleBackColor = true;
			// 
			// chbMenuId
			// 
			this.chbMenuId.AutoSize = true;
			this.chbMenuId.Location = new System.Drawing.Point(12, 35);
			this.chbMenuId.Name = "chbMenuId";
			this.chbMenuId.Size = new System.Drawing.Size(67, 17);
			this.chbMenuId.TabIndex = 3;
			this.chbMenuId.Text = "Menu ID";
			this.tt.SetToolTip(this.chbMenuId, "Adds the menu item in the context menu to show the menu ID and location.EPLAN "+
        "restart required.");
			this.chbMenuId.UseVisualStyleBackColor = true;
			// 
			// chbInplaceEditing
			// 
			this.chbInplaceEditing.AutoSize = true;
			this.chbInplaceEditing.Location = new System.Drawing.Point(12, 58);
			this.chbInplaceEditing.Name = "chbInplaceEditing";
			this.chbInplaceEditing.Size = new System.Drawing.Size(95, 17);
			this.chbInplaceEditing.TabIndex = 4;
			this.chbInplaceEditing.Text = "Inplace editing";
			this.tt.SetToolTip(this.chbInplaceEditing, "Direct editing, switching of the properties with the \"TAB\"- key.");
			this.chbInplaceEditing.UseVisualStyleBackColor = true;
			// 
			// chbDontChangeSourceText
			// 
			this.chbDontChangeSourceText.AutoSize = true;
			this.chbDontChangeSourceText.Location = new System.Drawing.Point(12, 81);
			this.chbDontChangeSourceText.Name = "chbDontChangeSourceText";
			this.chbDontChangeSourceText.Size = new System.Drawing.Size(141, 17);
			this.chbDontChangeSourceText.TabIndex = 5;
			this.chbDontChangeSourceText.Text = "Dont Change Source Text";
			this.tt.SetToolTip(this.chbDontChangeSourceText, "Disable capital letter conversion for source language.");
			this.chbDontChangeSourceText.UseVisualStyleBackColor = true;
			// 
			// chbRemoveSelection
			// 
			this.chbRemoveSelection.AutoSize = true;
			this.chbRemoveSelection.Location = new System.Drawing.Point(12, 104);
			this.chbRemoveSelection.Name = "chbRemoveSelection";
			this.chbRemoveSelection.Size = new System.Drawing.Size(110, 17);
			this.chbRemoveSelection.TabIndex = 6;
			this.chbRemoveSelection.Text = "Remove Selection";
			this.tt.SetToolTip(this.chbRemoveSelection, "Controls whether objects should be deselected after an interaction.");
			this.chbRemoveSelection.UseVisualStyleBackColor = true;
			// 
			// chbDebugScripts
			// 
			this.chbDebugScripts.AutoSize = true;
			this.chbDebugScripts.Location = new System.Drawing.Point(12, 127);
			this.chbDebugScripts.Name = "chbDebugScripts";
			this.chbDebugScripts.Size = new System.Drawing.Size(90, 17);
			this.chbDebugScripts.TabIndex = 7;
			this.chbDebugScripts.Text = "Debug Scripts";
			this.tt.SetToolTip(this.chbDebugScripts, "Controls whether objects should be deselected after an interaction.");
			this.chbDebugScripts.UseVisualStyleBackColor = true;
			// 
			// ExtendedSettings
			// 
			this.AcceptButton = this.btnOk;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(269, 196);
			this.Controls.Add(this.chbDebugScripts);
			this.Controls.Add(this.chbRemoveSelection);
			this.Controls.Add(this.chbDontChangeSourceText);
			this.Controls.Add(this.chbInplaceEditing);
			this.Controls.Add(this.chbMenuId);
			this.Controls.Add(this.chbCopySettingsPath);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.btnCancel);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ExtendedSettings";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Extended Settings";
			this.Load += new System.EventHandler(this.frmExtendedSettings_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        public ExtendedSettings()
        {
            InitializeComponent();
        }

        #endregion
        // ReSharper restore All

        [DeclareAction("ExtendedSettings")]
        public void Function(string name)
        {
            if (!SettingsParser.Ok(name))
            {
                return;
            }

            ExtendedSettings frm = new ExtendedSettings();
            frm.ShowDialog();
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void frmExtendedSettings_Load(object sender, System.EventArgs e)
        {
            // CopySettingsPath
            chbCopySettingsPath.Checked = settings.GetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowExtended", 0);

            // MenuId
            chbMenuId.Checked = settings.GetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowIdentifier", 0);

            // InplaceEditing
            chbInplaceEditing.Checked = settings.GetBoolSetting("USER.EnfMVC.Debug.InplaceEditingShowAllProperties", 0);

            // DontChangeSourceText
            chbDontChangeSourceText.Checked = settings.GetBoolSetting("USER.TRANSLATEGUI.DontChangeSourceText", 0);

            // RemoveSelection
            chbRemoveSelection.Checked = settings.GetBoolSetting("USER.GedViewer.RemoveSelection", 0);

			// DebugScripts
			chbDebugScripts.Checked = settings.GetBoolSetting("USER.EplanEplApiScriptLog.DebugScripts", 0);
		}

        private void btnOk_Click(object sender, System.EventArgs e)
        {
            // CopySettingsPath
            settings.SetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowExtended", chbCopySettingsPath.Checked, 0);

            // MenuId
            settings.SetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowIdentifier", chbMenuId.Checked, 0);

            // InplaceEditing
            settings.SetBoolSetting("USER.EnfMVC.Debug.InplaceEditingShowAllProperties", chbInplaceEditing.Checked, 0);

            // DontChangeSourceText
            settings.SetBoolSetting("USER.TRANSLATEGUI.DontChangeSourceText", chbDontChangeSourceText.Checked, 0);

            // RemoveSelection
            settings.SetBoolSetting("USER.GedViewer.RemoveSelection", chbRemoveSelection.Checked, 0);

			// RemoveSelection
			settings.SetBoolSetting("USER.EplanEplApiScriptLog.DebugScripts", chbDebugScripts.Checked, 0);

			Close();
        }

    }

    public static class SettingsParser
    {
        public static bool Ok(string name)
        {
            string value = GetText(name, "");
            if (string.IsNullOrEmpty(value))
            {
                MessageBox.Show(
                    "License could not be verified." + Environment.NewLine +
                    "The process was aborted. "," Warning - SettingsParser ",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else
            {
                return true;
            }
        }

        public static bool? GetBool(string name, string parameter)
        {
            string value = GetText(name, parameter);
            if (value.ToUpper().Equals("TRUE"))
            {
                return true;
            }
            else if (value.ToUpper().Equals("FALSE"))
            {
                return false;
            }
            else
            {
                return null;
            }
        }

        public static string GetCompanyName()
        {
            string value = GetText("", "COMPANYNAME");
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            else
            {
                return value;
            }
        }

        public static List<string> GetList(string name, string parameter)
        {
            string value = GetText(name, parameter);
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            else
            {
                return value.Split('|').ToList();
            }
        }

        public static string GetPath(string name, string parameter)
        {
            string value = GetText(name, parameter);
            try
            {
                if (value.Contains("$(PROJECTPATH_SUB)"))
                {
                    string directory =
                        Path.GetDirectoryName(PathMap.SubstitutePath("$(PROJECTPATH)"));

                    var path2 = value.Substring(value.IndexOf("$(PROJECTPATH_SUB)") + 18,
                        value.Length - 18);
                    if (directory != null)
                    {
                        return directory + path2;
                    }
                }
                else
                {
                    value = PathMap.SubstitutePath(value);
                    return value;
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        public static string GetPathFile(string name, string parameter)
        {
            string value = GetText(name, parameter);
            try
            {
                if (value.Contains("$(PROJECTPATH_SUB)"))
                {
                    string directory =
                        Path.GetDirectoryName(PathMap.SubstitutePath("$(PROJECTPATH)"));

                    var path2 = value.Substring(value.IndexOf("$(PROJECTPATH_SUB)") + 18,
                        value.Length - 18);
                    if (directory != null)
                    {
                        return directory + path2;
                    }
                }
                else
                {
                    value = PathMap.SubstitutePath(value);
                    return value;
                }

            }
            catch
            {
                return null;
            }
            return null;
        }

        public static string GetText(string name, string parameter)
        {
            string value = null;
            ActionCallingContext actionCallingContext = new ActionCallingContext();
            actionCallingContext.AddParameter("name", name);
            actionCallingContext.AddParameter("parameter", parameter);
            new CommandLineInterpreter().Execute("ScriptSettings", actionCallingContext);
            actionCallingContext.GetParameter("setting", ref value);

            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            else
            {
                return value;
            }
        }
    }
}

