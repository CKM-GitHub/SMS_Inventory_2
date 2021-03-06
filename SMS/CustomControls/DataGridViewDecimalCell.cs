using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SMS.CustomControls
{
    public class DataGridViewDecimalCell : DataGridViewTextBoxCell
    {
        //コンストラクタ
        public DataGridViewDecimalCell()
        {
        }

        //編集コントロールを初期化する
        //編集コントロールは別のセルや列でも使いまわされるため、初期化の必要がある
        public override void InitializeEditingControl(
            int rowIndex, object initialFormattedValue,
            DataGridViewCellStyle dataGridViewCellStyle)
            {
                base.InitializeEditingControl(rowIndex,
                initialFormattedValue, dataGridViewCellStyle);

                //編集コントロールの取得
                DataGridViewDecimalControl decimalBox =
                this.DataGridView.EditingControl as
                DataGridViewDecimalControl;
                if (decimalBox != null)
                {
                    //Textを設定
                    string decimalText = initialFormattedValue as string;
                    decimalBox.Text = decimalText != null ? decimalText : "";
                    //カスタム列のプロパティを反映させる
                    DataGridViewDecimalColumn column =
                    this.OwningColumn as DataGridViewDecimalColumn;
                    if (column != null)
                    {
                        decimalBox.DecimalPlace = column.DecimalPlace;
                        decimalBox.MaxLength = column.MaxInputLength;
                        decimalBox.UseThousandSeperator = column.UseThousandSeparator;
                        decimalBox.UseMinus = column.UseMinus;
                    }
                }
            }

        //編集コントロールの型を指定する
        public override Type EditType
        {
            get
            {
                return typeof(DataGridViewDecimalControl);
            }
        }

        //セルの値のデータ型を指定する
        //ここでは、Object型とする
        //基本クラスと同じなので、オーバーライドの必要なし
        public override Type ValueType
        {
            get
            {
                return typeof(object);
            }
        }

        //新しいレコード行のセルの既定値を指定する
        public override object DefaultNewRowValue
        {
            get
            {
                return base.DefaultNewRowValue;
            }
        }
    }
}
