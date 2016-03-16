using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace YUBI_TOOL.Model
{
    public interface IValidatableNestedObject
    {
        bool TryValidateNestedObject(ICollection<ValidationResult> validationResults);
        bool IsValid { get; }
    }
}
