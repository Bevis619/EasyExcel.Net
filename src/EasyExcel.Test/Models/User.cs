using EasyExcel.Export;
using System;
using System.Collections.Generic;

namespace EasyExcel.Test.Models
{
    //[EESheet("user data")]
    public class User
    {
        [EEHeader("编号")]
        public long Id { get; set; }

        [EEHeader("姓名")]
        public string Name { get; set; }

        [EEHeader("性别")]
        public string Sex { get; set; }

        [EEHeader("年龄")]
        public int Age => (DateTime.Now.Year - Birthday.Year) + 1;

        [EEHeader("生日")]
        public DateTime Birthday { get; set; }

        public static List<User> MakeUsers()
        {
            return new List<User>
            {
                new User{ Id = 1L,Name = "Bob", Sex = "男" , Birthday = new DateTime(1992,01,19) },
                new User{ Id = 2L,Name = "Tom", Sex = "男" , Birthday = new DateTime(1993,05,19) },
                new User{ Id = 3L,Name = "Marry", Sex = "女" , Birthday = new DateTime(1995,05,1) },
                new User{ Id = 4L,Name = "Jim", Sex = "男" , Birthday = new DateTime(1992,01,23) },
                new User{ Id = 4L,Name = "Sofafa", Sex = "女" , Birthday = new DateTime(1994,11,19) },
                new User{ Id = 5L,Name = "Joe", Sex = "男" , Birthday = new DateTime(1992,05,19) },
            };
        }
    }
}
