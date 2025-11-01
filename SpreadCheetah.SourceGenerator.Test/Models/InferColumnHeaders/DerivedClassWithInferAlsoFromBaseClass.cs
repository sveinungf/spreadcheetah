using SpreadCheetah.SourceGeneration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InheritColumns]
[InferColumnHeaders(typeof(ColumnHeaders), Suffix = "_Header")]
public class DerivedClassWithInferAlsoFromBaseClass : BaseClassWithInfer
{
    public required string Year { get; init; }
}
