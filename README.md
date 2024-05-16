# ucPrinterComboEx
Extended printer combobox control

![image](https://github.com/fafalone/ucPrinterComboEx/assets/7834493/f7a23f9d-d0c0-4a76-9cd6-6843e9be9d00) ![image](https://github.com/fafalone/ucPrinterComboEx/assets/7834493/871bf612-3985-47e6-9780-8d671553efc5)

ucPrinterComboEx is a simple control to select a printer, using a ListView to substitute for a normal dropdown to display the large icon/2-line display that manyu other selection dialogs use. There's a selection change event, methods for accessing the full collection, and some additional information like the 'Model' field from the Printers folder available, as well as showing the default printer in bold text. Krool's IPAO techniques are used to provide basic keyboard support. The dropdown behaves like a real combobox, including the slide animation (if enabled for regular combo controls), sliding up instead if there's not enough room on the bottom, extending beyond the form if needed, and fine control over sizing options.

Project is available as a VB6 UserControl and twinBASIC version. These versions are slightly different; the VB6 version has a set of 32bit-only declares that are picked up in twinBASIC by my WinDevLib project. Some changes in qualifying types are made to support using oleexp and OLEGuids in VB6 as well. But the codebase is 99% the same.

TestUCPC.twinproj - twinBASIC tbcontrol version on a test form.

ucPrinterComboEx.twinproj - Project configured to build as OCX.

ucPrinterComboExPackage.twinproj - Project configured to build as .twinpack

ucPrinterComboEx.twinpack - twinBASIC Package form suitable to import into tB projects.

The VB6 folder contains all files for that version, except for the associated typelibs oleexp.tlb and OLEGuids.tlb, which are available from:\
[[VB6] Modern Shell Interface Type Library - oleexp.tlb](http://www.vbforums.com/showthread.php?786079)\
[CommonControls (Replacement of the MS common controls)](https://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls))

---

Of course, hat tip to wqweto, I was wrong-- this project originally started thinking I could do that two line display easier with a ListView in tile mode than his owner draw combo method, turns out that no, that would have been easier and worked slightly better :D
