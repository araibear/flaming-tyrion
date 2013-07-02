using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassLibrary1
{

    public class Class1
    {

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
