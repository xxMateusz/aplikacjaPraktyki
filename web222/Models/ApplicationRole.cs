using Microsoft.AspNetCore.Identity;

namespace web222.Models;

public class ApplicationRole : IdentityRole
{
    public ICollection<ApplicationUserRole> UserRoles { get; set; }
}
