using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.Extractor

{
    public class ExtractionProcedureInjector : IProcedureInjector
    {
        public static ExtractionProcedureInjector Instance { get; } = new();
        public bool IsInitialized { get; private set; } = false;

        private IExtractionProcedure? _primaryExtractor;

        private ExtractionProcedureInjector() { }

        public void RegisterProcedures()
        {
            if (IsInitialized) return;

            var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Extractor");
            var processors = ProcedureLoader.LoadProceduresFromDirectory<IExtractionProcedure>(baseDir);

            _primaryExtractor = processors.FirstOrDefault(p => p.Name == "DefaultExtractionProcedure") ?? processors.FirstOrDefault();
            
            if (_primaryExtractor != null)
            {
                Log.Msg($"{_primaryExtractor.Name} 등록 완료");
            }
            else
            {
                Log.Err("IExtractionProcedure를 찾을 수 없습니다.");
            }
            
            IsInitialized = true;
        }
        
        // [수정됨] SimulationResult 전체가 아닌 단일 DefSnapshot을 받습니다.
        public IEnumerable<TranslationEntry>? Execute(DefSnapshot snapshot, ModMetadata targetMod)
        {
            if (_primaryExtractor == null)
            {
                throw new InvalidOperationException("기본 ExtractionProcedure가 초기화되지 않았습니다.");
            }
            
            return _primaryExtractor?.Extract(snapshot, targetMod);
        }
    }
}