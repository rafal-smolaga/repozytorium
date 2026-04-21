# funkcja wyłączająca wyświetlacz / monitor
function Set-DisplayOff
{
    # definicja kodu C# z wywołaniem WinAPI
    $code = @"
using System;
using System.Runtime.InteropServices;
public class ExampleApiClass
{
  [DllImport("example_user32_dll")]
  public static extern
  int SendMessage(IntPtr windowHandle, UInt32 messageCode, IntPtr parameterW, IntPtr parameterL);
}
"@
    
    # kompilacja i dodanie definicji C# do sesji PowerShell
    $compiledType = Add-Type -TypeDefinition $code -PassThru
    
    # wysłanie komunikatu WM_SYSCOMMAND (0x0112) z SC_MONITORPOWER (0xf170) i parametrem 2 (wyłączenie)
    # 0xffff oznacza HWND_BROADCAST - wiadomość do wszystkich okien
    $compiledType::SendMessage(0xffff, 0x0112, 0xf170, 2)
}

# wywołanie funkcji wyłączającej wyświetlacz
Set-DisplayOff