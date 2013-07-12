using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassLibrary1
{
    /* ----------------------------------------------------------------------
     * クラス：なんもしてないクラスです。
     * メソッド：TestMeth(String str1)　引数がいるのに使っていません（汗）
      --------------------------------------------------------------------*/
    public class Class1
    {
        /*-----------------------------------------------------------------------
         * 機能：a+b+cを計算します。
         * 引数：String str1(なんもいえねえ。)
         * 戻り値：６
         * ---------------------------------------------------------------------*/
        public String TestMeth(String str1)
        {
            int a, b, c;
            
            a=1;b=2;c=3;
            String str = "これは" + this.ToString() + "です。";
            str1 = (a + b + c).ToString();
            return str1;
        }


    }
}
