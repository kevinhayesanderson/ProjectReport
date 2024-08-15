namespace Models
{
    public class PunchMovementData(EmployeePunchData[] employeePunchDatas)
    {
        public int Length => employeePunchDatas.Length;

        public ref EmployeePunchData this[int index]
        {
            get
            {
                if (index < 0 || index >= Length)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }
                return ref employeePunchDatas[index];
            }
        }
    }
}