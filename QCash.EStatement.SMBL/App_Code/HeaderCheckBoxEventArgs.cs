using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Infragistics.Win.UltraWinGrid;
using System.Windows.Forms;

namespace QCash.EStatement
{
    // EventArgs used for the HeaderCheckBoxClicked event. This event has to pass in the CheckState and the ColumnHeader
    public class HeaderCheckBoxEventArgs : EventArgs
    {
        private Infragistics.Win.UltraWinGrid.ColumnHeader mvarColumnHeader;
        private CheckState mvarCheckState;
        private RowsCollection mvarRowsCollection;

        public HeaderCheckBoxEventArgs(Infragistics.Win.UltraWinGrid.ColumnHeader hdrColumnHeader, CheckState chkCheckState, RowsCollection Rows)
        {
            mvarColumnHeader = hdrColumnHeader;
            mvarCheckState = chkCheckState;
            mvarRowsCollection = Rows;
        }

        // Expose the rows collection for the specific row island that the header belongs to
        public RowsCollection Rows
        {
            get
            {
                return mvarRowsCollection;
            }
        }

        public Infragistics.Win.UltraWinGrid.ColumnHeader Header
        {
            get
            {
                return mvarColumnHeader;
            }
        }

        public CheckState CurrentCheckState
        {
            get
            {
                return mvarCheckState;
            }
            set
            {
                mvarCheckState = value;
            }
        }
    }			
}
