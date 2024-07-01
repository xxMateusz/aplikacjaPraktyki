using Microsoft.AspNetCore.Identity;

namespace web222.Models;

public class ApplicationUser : IdentityUser
{
 public ICollection<ApplicationUserRole> UserRoles { get; set; }
}