namespace WordDemo.Enums
{
    public enum OperationTypeEnum
    {
        /// <summary>
        /// 替换值
        /// </summary>
        ReplaceText = 0,

        /// <summary>
        /// 输出错误信息
        /// </summary>
        ConsoleError = 1,

        /// <summary>
        /// 输出错误信息且改变表头颜色
        /// </summary>
        ChangeColor = 2,

        /// <summary>
        /// 不操作
        /// </summary>
        NotOperation = 3

    }
}
