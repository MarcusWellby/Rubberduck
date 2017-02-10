using System;
using System.Collections;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ParameterDeclaration : Declaration, IInspectionResultTarget<ParameterDeclaration>
    {
        private readonly bool _isOptional;
        private readonly bool _isByRef;
        private readonly bool _isImplicitByRef;

        /// <summary>
        /// Creates a new built-in parameter declaration.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration, 
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            bool isOptional, 
            bool isByRef, 
            bool isArray = false, 
            bool isParamArray = false)
            : base(
                  qualifiedName, 
                  parentDeclaration, 
                  parentDeclaration, 
                  asTypeName,
                  typeHint,
                  false, 
                  false, 
                  Accessibility.Implicit,
                  DeclarationType.Parameter, 
                  null, 
                  Selection.Home,
                  isArray,
                  asTypeContext)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            _isImplicitByRef = false;
            IsParamArray = isParamArray;
        }

        /// <summary>
        /// Creates a new user declaration for a parameter.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration,
            ParserRuleContext context, 
            Selection selection, 
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            bool isOptional,
            bool isByRef,
            bool isArray = false, 
            bool isParamArray = false,
            bool isBuiltIn = false)
            : base(
                  qualifiedName, 
                  parentDeclaration, 
                  parentDeclaration,
                  asTypeName,
                  typeHint,
                  false, 
                  false, 
                  Accessibility.Implicit,
                  DeclarationType.Parameter, 
                  context, 
                  selection,
                  isArray,
                  asTypeContext,
                  isBuiltIn)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            _isImplicitByRef = isByRef && (context == null || ((VBAParser.ArgContext) context).BYREF() == null);
            IsParamArray = isParamArray;
        }

        public ParameterDeclaration(ComParameter parameter, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(parameter.Name),
                parent,
                parameter.TypeName,
                null,
                null,
                parameter.IsOptional,
                parameter.IsByRef,
                parameter.IsArray,
                parameter.IsParamArray)
        { }
             
        public bool IsOptional { get { return _isOptional; } }
        public bool IsByRef { get { return _isByRef; } }
        public bool IsImplicitByRef { get { return _isImplicitByRef; } }
        public bool IsParamArray { get; set; }

        private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

        #region ICollection<IInspectionResult>
        public IEnumerator<IInspectionResult> GetEnumerator()
        {
            return _inspectionTarget.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IInspectionResult item)
        {
            _inspectionTarget.Add(item);
        }

        public void Clear()
        {
            _inspectionTarget.Clear();
        }

        public bool Contains(IInspectionResult item)
        {
            return _inspectionTarget.Contains(item);
        }

        public void CopyTo(IInspectionResult[] array, int arrayIndex)
        {
            _inspectionTarget.CopyTo(array, arrayIndex);
        }

        [Obsolete("Throws NotSupportedException. Use Clear() method.")]
        public bool Remove(IInspectionResult item)
        {
            return false;
        }

        public int Count { get { return _inspectionTarget.Count; } }

        public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
        #endregion

        public ParameterDeclaration Target { get { return this; } }
    }
}
