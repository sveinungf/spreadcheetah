using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Collections.Generic;

namespace SpreadCheetah.SourceGenerator
{
    internal class SyntaxReceiver : ISyntaxReceiver
    {
        public List<SyntaxNode> ArgumentsToValidate { get; } = new();

        public void OnVisitSyntaxNode(SyntaxNode syntaxNode)
        {
            if (syntaxNode is InvocationExpressionSyntax
                {
                    ArgumentList:
                    {
                        Arguments:
                        {
                            Count: <= 3
                        } arguments
                    },
                    Expression: MemberAccessExpressionSyntax
                    {
                        Name:
                        {
                            Identifier:
                            {
                                ValueText: "AddAsRowAsync"
                            }
                        }
                    }
                })
            {
                ArgumentsToValidate.Add(arguments.First().Expression);
            }
        }
    }
}
