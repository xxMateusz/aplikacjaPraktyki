using Microsoft.AspNetCore.Identity;

namespace web222.Models;

public class ApplicationUserRole : IdentityUserRole<string>
{
    public ApplicationUser User { get; set; }
    public ApplicationRole Role { get; set; }
}