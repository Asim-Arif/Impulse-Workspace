using System;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Impulse.Services.Notifications;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace Impulse.Controllers
{
    [Route("api/push")]
    [ApiController]
    public class PushSubscriptionController : ControllerBase
    {
        private readonly IPushNotificationDataAccess _pushData;
        private readonly IWebPushNotificationService _webPushService;
        private readonly IUserDataAccess _userDataAccess;
        private readonly ILogger<PushSubscriptionController> _logger;

        public PushSubscriptionController(
            IPushNotificationDataAccess pushData,
            IWebPushNotificationService webPushService,
            IUserDataAccess userDataAccess,
            ILogger<PushSubscriptionController> logger)
        {
            _pushData = pushData;
            _webPushService = webPushService;
            _userDataAccess = userDataAccess;
            _logger = logger;
        }

        [HttpGet("vapid-public-key")]
        public IActionResult GetVapidPublicKey()
        {
            var publicKey = _webPushService.GetPublicKey();
            return Ok(new { publicKey });
        }

        [HttpPost("subscribe")]
        public async Task<IActionResult> Subscribe([FromBody] PushSubscriptionDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Endpoint) || dto.Keys == null)
            {
                return BadRequest("Invalid push subscription payload.");
            }

            try
            {
                var authenticatedUserName = User?.Identity?.Name
                    ?? User?.FindFirstValue(ClaimTypes.Name)
                    ?? User?.FindFirstValue(ClaimTypes.NameIdentifier);

                var userName = !string.IsNullOrWhiteSpace(authenticatedUserName)
                    ? authenticatedUserName
                    : (!string.IsNullOrWhiteSpace(dto.UserName) ? dto.UserName : dto.UserId);

                if (string.IsNullOrWhiteSpace(userName))
                {
                    return BadRequest("User identity is required to register push subscription.");
                }

                int? parsedUserId = null;
                if (!string.IsNullOrWhiteSpace(dto.UserId) && int.TryParse(dto.UserId, out var uid))
                {
                    parsedUserId = uid;
                }
                if (!parsedUserId.HasValue && !string.IsNullOrWhiteSpace(userName))
                {
                    try
                    {
                        var userObj = await _userDataAccess.GetUserByUserNameAsync(userName);
                        if (userObj != null && userObj.UserID > 0)
                        {
                            parsedUserId = userObj.UserID;
                        }
                    }
                    catch
                    {
                        // Fallback if user lookup is not critical
                    }
                }

                var userAgent = Request.Headers.UserAgent.ToString();

                await _pushData.UpsertSubscriptionAsync(
                    userName: userName,
                    userId: parsedUserId,
                    endpoint: dto.Endpoint,
                    p256dh: dto.Keys.P256dh,
                    auth: dto.Keys.Auth,
                    deviceName: userAgent
                );
                _logger.LogInformation("Successfully registered/updated push subscription for user {UserName} (ID: {UserId}) on device {Endpoint}", userName, parsedUserId, dto.Endpoint);

                return Ok(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to save push subscription.");
                return StatusCode(500, new { success = false, message = "Failed to save push subscription." });
            }
        }

        [HttpPost("unsubscribe")]
        public async Task<IActionResult> Unsubscribe([FromBody] UnsubscribeDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Endpoint))
            {
                return BadRequest("Endpoint is required.");
            }

            try
            {
                await _pushData.DeactivateSubscriptionAsync(dto.Endpoint);
                return Ok(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to deactivate push subscription.");
                return StatusCode(500, new { success = false, message = "Failed to deactivate push subscription." });
            }
        }

        public class UnsubscribeDto
        {
            public string Endpoint { get; set; } = string.Empty;
        }
    }
}
