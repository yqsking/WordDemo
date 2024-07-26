

namespace WordDemo.Dtos
{
    public interface IProgressModel
    {
        string Title { get; set; }
        int CurrentStep { get; set; }
        int TotalSteps { get; set; }
        string Message { get; set; }
    }
}
