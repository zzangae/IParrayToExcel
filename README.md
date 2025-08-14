# IParrayToExcel
엑셀에서 IP를 순서대로 정렬한다

## 1) 프로젝트 만들기 (.NET Framework 기반)
Visual Studio Code에서는 .NET Framework Class Library 템플릿이 기본 제공되지 않으니, 다음 중 하나를 선택해야 합니다.
- 방법 A (추천): Visual Studio Community에서 한 번만 .NET Framework Class Library 프로젝트를 만들고 이후 코딩·빌드는 VS Code에서 진행
- 방법 B: VS Code + MSBuild + .NET Framework SDK 설치 후 .csproj 직접 작성
예시 .csproj (Framework 4.8, COMVisible=true):

<pre>
&lt;Project Sdk="Microsoft.NET.Sdk">

  &lt;PropertyGroup>
    &lt;TargetFramework>net48&lt;/TargetFramework>
    &lt;OutputType>Library&lt;/OutputType>
    &lt;GenerateAssemblyInfo>false&lt;/GenerateAssemblyInfo>
    &lt;RegisterForComInterop>true&lt;/RegisterForComInterop>
    &lt;PlatformTarget>x86&lt;/PlatformTarget> &lt;!-- Office 비트수에 맞추기 -->
  &lt;/PropertyGroup>

  &lt;ItemGroup>
    &lt;Reference Include="System" />
    &lt;Reference Include="System.Net" />
    &lt;Reference Include="System.Runtime.InteropServices" />
  &lt;/ItemGroup>

&lt;/Project>
</pre>
* PlatformTarget은 Office 비트수와 동일해야 함 (32비트 Office → x86, 64비트 Office → x64)

## 2) C# 코드 작성
제가 위에서 드린 IpToolsLib 코드 파일을 프로젝트 폴더에 저장 (IpTools.cs).

## 3) 빌드
VS Code 터미널에서:
<pre>
msbuild IpToolsLib.csproj /p:Configuration=Release
</pre>
* 결과 DLL은 bin\Release 폴더에 생성됨.

## 4) COM 등록 (regasm 사용)
관리자 권한 PowerShell/명령 프롬프트에서:
64비트 Office:
<pre>
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm "경로\IpToolsLib.dll" /codebase /tlb:IpToolsLib.tlb
</pre>
32비트 Office:
<pre>
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm "경로\IpToolsLib.dll" /codebase /tlb:IpToolsLib.tlb
</pre>

## 5) VBA에서 호출
#### 조기 바인딩
VBA → 도구 → 참조 → IpToolsLib 체크
<pre>
Sub Test()
    Dim tool As IpToolsLib.IpTools
    Set tool = New IpToolsLib.IpTools
    MsgBox tool.IsValid("192.168.0.1")
End Sub
</pre>
#### 후기 바인딩
<pre>
Sub TestLate()
    Dim tool As Object
    Set tool = CreateObject("IpTools.IpTools")
    MsgBox tool.IsValid("192.168.0.1")
End Sub
</pre>

## 6) VS Code 만 쓸 때 유의점
- .NET Framework SDK와 MSBuild 설치 필요 (Visual Studio Build Tools 설치 시 포함)
- .csproj 직접 편집해야 COM 등록 관련 설정 가능
- 빌드 후 COM 등록은 regasm으로 수동 진행
- Office 비트수와 DLL 플랫폼 일치 필수
- 난독화(소스 보호)는 추가로 dotfuscator나 ConfuserEx 같은 툴 사용 가능
