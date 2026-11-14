using System.Runtime.CompilerServices;
using System.Xml.Linq;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator;

/// <summary>DefTreeSimulator에서 동작하는 프로시저 인젝터의 구현체입니다.</summary>
public class DefTreeProcedureInjector : IProcedureInjector
{
    public static DefTreeProcedureInjector Instance { get; } = new();

    private readonly List<(IDefTreeProcedure Proc, bool IsEnabled)> _procedures = new();
    private readonly object _syncLock = new();
    
    public bool IsInitialized { get; private set; } = false;

    private DefTreeProcedureInjector() { }

    public void RegisterProcedures()
    {
        if (IsInitialized) return;

        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "DefTree");
        var loadedProcs = ProcedureLoader.LoadProceduresFromDirectory<IDefTreeProcedure>(baseDir);

        lock (_syncLock)
        {
            _procedures.Clear();
            foreach (var proc in loadedProcs)
            {
                if (!_procedures.Any(p => p.Proc.Name == proc.Name))
                {
                    _procedures.Add((proc, false));
                    Log.Good($"{proc.Name} 등록 완료.", true);
                }
            }

            IsInitialized = true;
        }

        SyncWithSettings();
    }

    /// <summary>SettingManager.Current의 활성화 목록 및 실행 순서를 인젝터 내부 캐시와 동기화합니다.</summary>
    public void SyncWithSettings()
    {
        if (!IsInitialized)
        {
            RegisterProcedures();
            return;
        }

        lock (_syncLock)
        {
            var enabledNames = SettingManager.Current.EnabledDefTreeProcedures;
            var reordered = new List<(IDefTreeProcedure Proc, bool IsEnabled)>(_procedures.Count);
            var remaining = _procedures.Select(x => x.Proc).ToList();

            // 1. 설정에 저장된 순서대로 활성화된 프로시저 배치
            foreach (var name in enabledNames)
            {
                var match = remaining.FirstOrDefault(p => p.Name == name);
                if (match != null)
                {
                    reordered.Add((match, true));
                    remaining.Remove(match);
                }
            }

            // 2. 비활성화된 나머지 프로시저들도 인스턴스를 유지한 채 뒤에 보관
            foreach (var proc in remaining)
            {
                reordered.Add((proc, false));
            }

            _procedures.Clear();
            _procedures.AddRange(reordered);
        }
    }

    private void ExecuteInner(Action<IDefTreeProcedure> action, [CallerMemberName] string callerName = "")
    {
        List<(IDefTreeProcedure Proc, bool IsEnabled)> snapshotProcedures;
        lock (_syncLock) // 과한 방어. 아마 필요하진 않음.
        {
            snapshotProcedures = _procedures.ToList();
        }

        foreach (var (proc, isEnabled) in snapshotProcedures)
        {
            if (!isEnabled) continue;

            try
            {
                action(proc);
            }
            catch (Exception e)
            {
                Log.Err($"{proc.Name}/{callerName} 실행 실패\n\t{e.Message}");
            }
        }
    }
    
    /// <summary>사전 해석 트리(PrePiledTree)가 로드된 직후 등록된 프로시저들을 실행합니다.</summary>
    public static void ExecutePostPrePile(XDocument prePiledTree)
        => Instance.ExecuteInner(proc => proc.PostPrePile(prePiledTree));
    
    /// <summary>Def 트리 병합 및 상속 처리가 끝난 직후 등록된 프로시저들을 실행합니다.</summary>
    public static void ExecutePostCreateTree(DefSnapshot snapshot)
        => Instance.ExecuteInner(proc => proc.PostCreateTree(snapshot));
}