using System;
using System.Collections.Generic;
using System.Text;

namespace SPUtil.Infrastructure
{
    public enum AuthResult
    {
        Success,
        InvalidCredentials, // 401: Плохой пароль
        AccessDenied,      // 403: Нет прав
        SiteNotFound,      // 404: Ошиблись в URL
        URLIsEmpty,
        Error              // Прочие ошибки
    }
}
