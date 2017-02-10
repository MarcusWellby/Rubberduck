using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IDeclarationWithParameter
    {
        IEnumerable<ParameterDeclaration> Parameters { get; }
        void AddParameter(ParameterDeclaration parameter);
    }
}
