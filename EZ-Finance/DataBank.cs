namespace EZ_Finance
{
    /*Future notes* 
    Add encapsulation for better access management. */ 

    class DCU_data
    {
        // *NOTE* // DCU User needs to be a double because DCU username is just a member number with no alpha characters in it. 
        public string User { get; set; }  
        public string Pass { get; set; }

        public string Date { get; set; }
        public string Description { get; set; }
        public string Deposit { get; set; }
        public string Withdrawl { get; set; }
        public string Balance { get; set; }

        public string Checking { get; set; }
        public string Savings { get; set; }
        public string Credit { get; set; }
    }


    class NET_data
    {
        public string User { get; set; }
        public string Pass { get; set; }

        public string Date { get; set; }
        public string Description { get; set; }
        public string Amount { get; set; }
        public string Balance { get; set; }
        

        public string Checking { get; set; }
        public string Savings { get; set; }
    }

    class FileName_data
    {
        public string currentDate { get; set; }
        public string currentUser { get; set; }
        public string excelPath { get; set; }
    }

    class MailSender_data
    {
        public string senderMailAddress { get; set; }
        public string senderMailPassword { get; set; }
    }

}

