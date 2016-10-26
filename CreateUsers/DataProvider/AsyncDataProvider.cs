using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CreateUsers.Model;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Windows.Documents;

namespace CreateUsers.DataProvider
{
    public static class AsyncDataProvider
    {
        // Получить пустое дерево подразделений домена
        public static ReadOnlyCollection<DomainTreeItem> GetDomainOUTree()
        {
            return new ReadOnlyCollection<DomainTreeItem>(new List<DomainTreeItem>().AsReadOnly());
        }
        // Получить дерево подразделений домена
        public static ReadOnlyCollection<DomainTreeItem> GetDomainOUTree(DirectoryEntry entry, ref string errorMsg)
        {
            List<DomainTreeItem> result = new List<DomainTreeItem>();
            DomainTreeItem rootNode = null;
            DomainTreeItem treeNode = null;
            DomainTreeItem curNode = null;
            string domain = "";
            string domainTag = "";
            string curTag = "";

            try
            {
                DirectorySearcher searchAD = new DirectorySearcher(entry);
                searchAD.Filter = "distinguishedName=*";
                SearchResult resultsAD = searchAD.FindOne();
                string domainDN = (string)resultsAD.Properties["distinguishedName"][0];

                DirectorySearcher dirSearcher = new DirectorySearcher(entry);
                dirSearcher.SearchScope = SearchScope.Subtree;
                dirSearcher.Filter = string.Format("(|(objectClass=organizationalUnit)(objectClass=organization)(cn=Users)(cn=Computers))");
                dirSearcher.PropertiesToLoad.Add("distinguishedName");
                dirSearcher.PropertiesToLoad.Add("l");
                SearchResultCollection searchResults = dirSearcher.FindAll();

                #region Создане корневого элемента дерева
                foreach (string dom in domainDN.Split(','))
                {
                    if (domain.Length > 0)
                    {
                        domain += "." + dom.Split('=')[1];
                        domainTag += "," + dom;
                    }
                    else
                    {
                        domain += dom.Split('=')[1];
                        domainTag += dom;
                    }
                }
                rootNode = new DomainTreeItem() { Title = domain, Description = domainTag, Image = @"/CreateUsers;component/Resources/ADServer.ico", IsSelected = true, IsExpanded = true };
                result.Add(rootNode);
                #endregion

                #region Обработка списка входных данных которые состоят из dn записей OU домена
                foreach (SearchResult sr in searchResults)
                {
                    // Разбиваем dn запись на элементы для дальнейшей обработки, в качестве разделителя используется запятая
                    string[] arrdata = ((string)sr.Properties["distinguishedName"][0]).Split(',');
                    // Получаем город dn записи
                    string city = (sr.Properties.Contains("l") ? (string)sr.Properties["l"][0] : "");
                    // Домен каждый раз получаем из dn записи
                    domain = "";
                    domainTag = "";

                    #region Получение домена из dn записи
                    foreach (string dom in arrdata)
                    {
                        if (dom.StartsWith("DC"))
                        {
                            if (domain.Length > 0)
                            {
                                domain += "." + dom.Split('=')[1];
                                domainTag += "," + dom;
                            }
                            else
                            {
                                domain += dom.Split('=')[1];
                                domainTag += dom;
                            }
                        }
                    }
                    #endregion

                    //Обнуляем дерево dn записи
                    treeNode = null;

                    #region Проверяем совпадает ли домен полученый из dn записи с доменом по умолчанию, совпадает - берем домен по умолчанию, не совпадает создаём дочерний домен
                    if (rootNode.Description == domainTag)
                    {
                        treeNode = rootNode;
                        curTag = rootNode.Description;
                    }
                    else
                    {
                        foreach (DomainTreeItem tvi in rootNode.Childs)
                        {
                            if (tvi.Description == domainTag)
                            {
                                treeNode = tvi;
                                curTag = tvi.Description;
                                break;
                            }
                        }
                        if (treeNode == null)
                        {
                            treeNode = new DomainTreeItem() { Title = domain, Description = domainTag, Image = @"/CreateUsers;component/Resources/ADServer.ico", IsSelected = true, IsExpanded = true };
                            rootNode.Childs.Add(treeNode);
                            curTag = domainTag;
                        }
                    }
                    #endregion

                    //Обнуляем текущий элемент дерева
                    curNode = null;

                    #region Добавляем текущую dn запись в дерево домена, если какой либо ветки нет в домене она будет создана
                    //Получем количество элементов масива dn записи
                    int i = arrdata.Length - 1;
                    while (i >= 0)
                    {
                        //Если елемент массива начинается с DC, это часть имени домена, переходим к следующему элементу
                        if (arrdata[i].StartsWith("DC"))
                        {
                            i--;
                            continue;
                        }

                        //формируем текущий тег записи
                        curTag = arrdata[i] + "," + curTag;
                        //Обнуляем значение переменной нового элемента дерева
                        DomainTreeItem newNode = null;


                        if (curNode != null)
                        {
                            //Если текущий элемент дерева не пуст, значит ищем новый элемент дерева в нём
                            foreach (DomainTreeItem tvi in curNode.Childs)
                            {
                                //Если в текущем элементе дерева нашли новый элемент, значит делаем найденный элемент новым
                                if (tvi.Description == curTag)
                                {
                                    newNode = tvi;
                                    break;
                                }
                            }
                            //Если новый элемент дерева пуст значит создаём его
                            if (newNode == null)
                            {
                                newNode = new DomainTreeItem() { Title = arrdata[i].Split('=')[1], Description = curTag, City = city, Image = @"/CreateUsers;component/Resources/ActiveDirectory.ico" };
                                curNode.Childs.Add(newNode);
                            }
                            //Делаем новый элемент текущим
                            curNode = newNode;
                        }
                        else
                        {
                            //Если текущий элемент дерева пуст, значит осущесвляем поиск нового элемента в корневом элементе дерева
                            foreach (DomainTreeItem tvi in treeNode.Childs)
                            {
                                //Если в корневом элементе дерева нашли новый элемент, значит делаем найденный элемент новым
                                if (tvi.Description == curTag)
                                {
                                    newNode = tvi;
                                    break;
                                }
                            }
                            //Если новый элемент дерева пуст значит создаём его
                            if (newNode == null)
                            {
                                newNode = new DomainTreeItem() { Title = arrdata[i].Split('=')[1], Description = curTag, City = city, Image = @"/CreateUsers;component/Resources/ActiveDirectory.ico" };
                                treeNode.Childs.Add(newNode);
                            }
                            //Делаем новый элемент текущим
                            curNode = newNode;
                        }
                        //переходим к следующему элементу массива
                        i--;
                    }
                    #endregion

                }
                #endregion
            }
            catch (Exception exp)
            {
                errorMsg = exp.ToString();
            }

            return result.AsReadOnly();
        }
        // Получить пустой список групп
        public static ReadOnlyCollection<Group> GetGroupItems()
        {
            return new ReadOnlyCollection<Group>(new List<Group>().AsReadOnly());
        }
        // Получить список всех групп в домене для выбора
        public static ReadOnlyCollection<Group> GetGroupForSelected(DirectoryEntry entry, ref string errorMsg)
        {
            List<Group> result = new List<Group>();
            try
            {
                DirectorySearcher dirSearcher = new DirectorySearcher(entry);
                dirSearcher.SearchScope = SearchScope.Subtree;
                dirSearcher.SizeLimit = 10000;
                dirSearcher.PageSize = 10000;
                dirSearcher.Filter = string.Format("objectClass=group");
                dirSearcher.PropertiesToLoad.Add("distinguishedName");
                dirSearcher.PropertiesToLoad.Add("sAMAccountName");
                dirSearcher.PropertiesToLoad.Add("description");
                SearchResultCollection searchResults = dirSearcher.FindAll();
                foreach (SearchResult sr in searchResults)
                {
                    result.Add(new Group
                    {
                        PlaceInAD = (string)sr.Properties["distinguishedName"][0],
                        Name = (string)sr.Properties["sAMAccountName"][0],
                        Description = (sr.Properties.Contains("description") ? (string)sr.Properties["description"][0] : ""),
                    });
                }
            }
            catch (Exception exp)
            {
                errorMsg = exp.ToString();
            }
            return result.AsReadOnly();
        }
    }
}
