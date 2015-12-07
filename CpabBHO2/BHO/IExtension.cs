using System;
using System.Runtime.InteropServices;

namespace CpabBHO2.BHO
{
    [
        ComVisible(true),
        Guid("0bda2434-16ae-4af5-be91-b11409482f86"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
        ]
    /* 扩展接口，用于注入的JS脚本回调BHO */
    public interface IExtension
    {
        [DispId(1)]
        void Foo(string s);
        [DispId(1)]
        void SaveCustomer(string s);
        [DispId(1)]
        void BlurCusNo(string s);
        [DispId(1)]
        void Mobile(string s);
    }
}
