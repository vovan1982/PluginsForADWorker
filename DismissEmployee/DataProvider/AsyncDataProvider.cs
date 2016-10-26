using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using DismissEmployee.Model;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Windows.Documents;

namespace DismissEmployee.DataProvider
{
    public static class AsyncDataProvider
    {
        #region Методы получения информации для вкладки "Пользователи"
        // Получить пустой список пользователей
        public static ReadOnlyCollection<User> GetItems()
        {
            return new ReadOnlyCollection<User>(new List<User>().AsReadOnly());
        }
        // Получить список пользователей для выбора в окне добавления пользователей в группу
        public static ReadOnlyCollection<User> GetUsersForSelected(string fieldForSearch, string searchStr, PrincipalContext context, DirectoryEntry entry, ref string errorMsg)
        {
            if (string.IsNullOrWhiteSpace(fieldForSearch) || string.IsNullOrWhiteSpace(searchStr))
            {
                return new ReadOnlyCollection<User>(new List<User>().AsReadOnly());
            }
            List<User> items = new List<User>();
            try
            {
                DirectorySearcher dirSearcher = new DirectorySearcher(entry);
                dirSearcher.SearchScope = SearchScope.Subtree;
                if (fieldForSearch == "Default")
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|(cn=*{0}*)(sn=*{0}*)(givenName=*{0}*)(sAMAccountName=*{0}*)))", searchStr.Trim());
                else
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|({1}=*{0}*)))", searchStr.Trim(), fieldForSearch);
                dirSearcher.PropertiesToLoad.Add("sAMAccountName");
                dirSearcher.PropertiesToLoad.Add("name");
                dirSearcher.PropertiesToLoad.Add("title");
                dirSearcher.PropertiesToLoad.Add("department");
                dirSearcher.PropertiesToLoad.Add("l");
                dirSearcher.PropertiesToLoad.Add("mobile");
                dirSearcher.PropertiesToLoad.Add("company");
                dirSearcher.PropertiesToLoad.Add("streetAddress");
                dirSearcher.PropertiesToLoad.Add("pwdLastSet");
                SearchResultCollection searchResults = dirSearcher.FindAll();
                foreach (SearchResult sr in searchResults)
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, (string)sr.Properties["sAMAccountName"][0]);
                    items.Add(new User
                    {
                        PlaceInAD = user.DistinguishedName,
                        Name = user.GivenName,
                        Surname = user.Surname,
                        DisplayName = user.DisplayName,
                        Login = (string)sr.Properties["sAMAccountName"][0],
                        NameInAD = (string)sr.Properties["name"][0],
                        Post = (sr.Properties.Contains("title") ? (string)sr.Properties["title"][0] : ""),
                        PhoneInt = user.VoiceTelephoneNumber,
                        PhoneMob = (sr.Properties.Contains("mobile") ? (string)sr.Properties["mobile"][0] : ""),
                        City = (sr.Properties.Contains("l") ? (string)sr.Properties["l"][0] : ""),
                        Department = (sr.Properties.Contains("department") ? (string)sr.Properties["department"][0] : ""),
                        Mail = user.EmailAddress,
                        Organization = (sr.Properties.Contains("company") ? (string)sr.Properties["company"][0] : ""),
                        Adress = (sr.Properties.Contains("streetAddress") ? (string)sr.Properties["streetAddress"][0] : ""),
                        AccountIsDisable = (user.Enabled == false ? true : false),
                    });
                }
            }
            catch (Exception exp)
            {
                errorMsg = exp.ToString();
            }
            return items.AsReadOnly();
        }
        #endregion
    }
}
