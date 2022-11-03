namespace NPOI_Demo.Models
{
    public class MissionAuto
    {
        public int Order { get; set; }
        public int Group { get; set; }
        public string CriteriaType { get; set; }
        public int CriteriaAmount { get; set; }
        public string RewardType { get; set; }
        public string RewardProvider { get; set; }
        public int RewardAmount { get; set; }
        public int PeriodInHours { get; set; }
    }
}