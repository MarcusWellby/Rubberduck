﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class PublicSubListener : IVBBaseListener, IExtensionListener<VBParser.SubStmtContext>
    {
        private readonly IList<VBParser.SubStmtContext> _members = new List<VBParser.SubStmtContext>();
        public IEnumerable<VBParser.SubStmtContext> Members { get { return _members; } }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            var visibility = context.Visibility();
            if (visibility == null || visibility.PUBLIC() != null)
            {
                _members.Add(context);
            }
        }
    }
}