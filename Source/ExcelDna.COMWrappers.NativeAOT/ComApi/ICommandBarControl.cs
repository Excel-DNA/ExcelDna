using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;
using Managed = Addin.Types.Managed;
using Marshalling = Addin.Types.Marshalling;

namespace Addin.ComApi;

[GeneratedComInterface]
[Guid("000C030E-0000-0000-C000-000000000046")]
public partial interface ICommandBarButton
{
}

[GeneratedComInterface]
[Guid("000C030A-0000-0000-C000-000000000046")]
public partial interface ICommandBarPopup
{
}

[GeneratedComInterface]
[Guid("000C030C-0000-0000-C000-000000000046")]
public partial interface ICommandBarComboBox
{
}

//[GeneratedComInterface]
//[Guid("000C0306-0000-0000-C000-000000000046")]
//public partial interface ICommandBarControls
//{
//    [PreserveSig]
//    object Add(object Type, object Id, object Parameter, object Before, object Temporary);
//}
