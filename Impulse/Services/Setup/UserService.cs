using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Setup
{
    public class UserService : IUserService
    {
        private readonly IUserDataAccess _userDataAccess;
        private readonly ILogger<UserService> _logger;

        public UserService(IUserDataAccess userDataAccess, ILogger<UserService> logger)
        {
            _userDataAccess = userDataAccess;
            _logger = logger;
        }

        public async Task<List<UserModel>> GetUsersAsync(string? search = null, bool? activeOnly = null)
        {
            return await _userDataAccess.GetUsersAsync(search, activeOnly);
        }

        public async Task<UserModel?> GetUserByIdAsync(int userId)
        {
            return await _userDataAccess.GetUserByIdAsync(userId);
        }

        public async Task<UserModel?> GetUserByUserNameAsync(string userName)
        {
            return await _userDataAccess.GetUserByUserNameAsync(userName);
        }

        public async Task<(bool Success, string Message, int UserId)> CreateUserAsync(UserModel user)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(user.UserName))
                {
                    return (false, "Username is required.", 0);
                }

                user.UserName = user.UserName.Trim();

                // Check duplicate username
                var existing = await _userDataAccess.GetUserByUserNameAsync(user.UserName);
                if (existing != null)
                {
                    return (false, $"A user with username '{user.UserName}' already exists.", 0);
                }

                var newId = await _userDataAccess.CreateUserAsync(user);
                return (true, "User created successfully.", newId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating user {UserName}", user.UserName);
                return (false, $"Failed to create user: {ex.Message}", 0);
            }
        }

        public async Task<(bool Success, string Message)> UpdateUserAsync(UserModel user)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(user.UserName))
                {
                    return (false, "Username is required.");
                }

                user.UserName = user.UserName.Trim();

                // Check duplicate username for other user
                var existing = await _userDataAccess.GetUserByUserNameAsync(user.UserName);
                if (existing != null && existing.UserID != user.UserID)
                {
                    return (false, $"Username '{user.UserName}' is already assigned to another account.");
                }

                var success = await _userDataAccess.UpdateUserAsync(user);
                return success ? (true, "User updated successfully.") : (false, "No changes were recorded.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating user ID {UserID}", user.UserID);
                return (false, $"Failed to update user: {ex.Message}");
            }
        }

        public async Task<bool> ToggleUserStatusAsync(int userId, bool inActive)
        {
            return await _userDataAccess.ToggleUserStatusAsync(userId, inActive);
        }

        public async Task<(bool Success, string Message)> ResetPasswordAsync(int userId, string newPassword)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(newPassword))
                {
                    return (false, "New password cannot be empty.");
                }

                var success = await _userDataAccess.ResetPasswordAsync(userId, newPassword.Trim());
                return success ? (true, "Password updated successfully.") : (false, "Failed to update password.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error resetting password for user {UserID}", userId);
                return (false, $"Error resetting password: {ex.Message}");
            }
        }

        public async Task<bool> DeleteUserAsync(int userId)
        {
            return await _userDataAccess.DeleteUserAsync(userId);
        }

        public async Task<(int TotalUsers, int ActiveUsers, int InactiveUsers)> GetUserStatisticsAsync()
        {
            var total = await _userDataAccess.GetTotalUsersCountAsync();
            var active = await _userDataAccess.GetActiveUsersCountAsync();
            var inactive = Math.Max(0, total - active);
            return (total, active, inactive);
        }

        public async Task<(bool Success, string Message, int NewUserId)> CopyUserAsync(int fromUserId, string newUserName, string? password = null, string? fullUserName = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(newUserName))
                {
                    return (false, "New username is required.", 0);
                }

                newUserName = newUserName.Trim();

                // Check duplicate username
                var existing = await _userDataAccess.GetUserByUserNameAsync(newUserName);
                if (existing != null)
                {
                    return (false, $"A user with username '{newUserName}' already exists.", 0);
                }

                var sourceUser = await _userDataAccess.GetUserByIdAsync(fromUserId);
                if (sourceUser == null)
                {
                    return (false, $"Source user ID #{fromUserId} could not be found.", 0);
                }

                var newUserId = await _userDataAccess.CopyUserAsync(fromUserId, newUserName, password?.Trim(), fullUserName?.Trim());
                return (true, $"User profile '{sourceUser.UserName}' successfully copied to '{newUserName}'.", newUserId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error copying user ID {FromUserId} to {NewUserName}", fromUserId, newUserName);
                return (false, $"Failed to copy user: {ex.Message}", 0);
            }
        }
    }
}
