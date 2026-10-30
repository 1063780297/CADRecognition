# libs — 第三方二进制依赖

本目录存放无法通过 NuGet 还原的第三方 DLL，随源码一起分发。项目通过相对路径 `..\libs\xxx.dll`
引用这里的文件，因此**克隆/解压后无需任何额外配置即可编译**。

## 清单

| 文件 | 版本 | 大小 | SHA256 | 用途 |
| --- | --- | --- | --- | --- |
| `CADImport.dll` | 16.1.0.65224 | 2105 KB | `A58803BF31F6EF9B95BE044D482D589709BF3A530ED80C630F71C295CE61EEA8` | DWG/DXF 图纸解析（SoftGold CAD.NET v16） |
| `HslCommunication.dll` | 12.7.0.0 | 4853 KB | 见下方校验命令 | PLC / Modbus TCP 通信 |

## 引用位置

- `CADRecognition/CADRecognition.csproj` → `<Reference Include="CADImport">`、`<Reference Include="HslCommunication">`
- `Demo/CadNetDemo.csproj` → `<Reference Include="CADImport">`

两者均设置了 `<Private>true</Private>`，编译后会复制到 `bin` 输出目录。

## 升级步骤

1. 用新版本的 DLL 覆盖本目录同名文件；
2. 更新上表的版本号与 SHA256；
3. 重新编译，确认 `bin\Debug\net48\` 下的 DLL 版本与预期一致。

## 授权提醒

- **CADImport.dll** 为 SoftGold 商业组件，随源码分发前请确认授权允许接收方使用；
- **HslCommunication.dll** 为授权版本，激活码硬编码在 `CADRecognition/HslCommunicationAppLicense.cs`，
  即分发源码等于一并交出该激活码，同样需要确认授权范围。

## 校验命令

```powershell
Get-FileHash -LiteralPath .\libs\*.dll -Algorithm SHA256
```
