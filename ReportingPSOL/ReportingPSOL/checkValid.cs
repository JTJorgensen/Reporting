using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportingPSOL
{
    class checkValid
    {
        public String error;

        private bool checkAccountSelection(String check)
        {
            bool result;

            if (check != null)
            {
                result = true;
            }
            else
            {
                result = false;
                error += "Please select a VALID account before continuing." + "\r\n";
            }

            return result;
        }//end checkAccountSelection()

        private bool checkAccountSelection(object check)
        {
            bool result;
            if (check != null)
            {
                result = checkAccountSelection(check.ToString());
            }
            else
            {
                result = false;
                error += "Please make an account selection before continuing." + "\r\n";
            }

            return result;
        }//end checkAccountSelection()

        private bool checkReportType(RadioButton rdoBilling, RadioButton rdoTicketing, RadioButton rdoTechnician)
        {
            bool result;

            if (rdoBilling.Checked)
            {
                result = true;
            }
            else if (rdoTicketing.Checked)
            {
                result = true;
            }
            else if (rdoTechnician.Checked)
            {
                result = true;
            }
            else
            {
                result = false;
                error += "Please select a report type before continuing." + "\r\n";
            }

            return result;
        }//end checkReportType()

        public bool checkAllForValid(RadioButton rdoBilling, RadioButton rdoTicketing, RadioButton rdoTechnician, object cmbAccount)
        {
            error = "";
            bool result;
            bool chkReport = checkReportType(rdoBilling, rdoTicketing, rdoTechnician);
            
            if (chkReport)
            {
                if (!rdoTechnician.Checked)
                {
                    result = checkAccountSelection(cmbAccount);
                }
                else if (rdoTechnician.Checked)
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            else
            {
                result = false;
            }

            return result;
        }//end checkAllForValid()
    }//end class
}//end namespace
