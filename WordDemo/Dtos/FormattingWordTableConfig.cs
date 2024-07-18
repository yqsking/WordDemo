
using WordDemo.Enums;

namespace WordDemo.Dtos
{
    public class FormattingWordTableConfig
    {
        /// <summary>
        /// 实线边框表格水平位置
        /// </summary>
        public HorizontalPositionTypeEnum SolidLineBorderTableHorizontalPositionType { get; set; }
        /// <summary>
        /// 实线边框表格垂直位置
        /// </summary>
        public VerticalPositionTypeEnum SolidLineBorderTableVerticalPositionTypeEnum { get; set; }

        /// <summary>
        /// 其他边框表格水平位置
        /// </summary>
        public HorizontalPositionTypeEnum OtherHorizontalPositionType { get; set; }
        /// <summary>
        /// 其他边框表格垂直位置
        /// </summary>
        public VerticalPositionTypeEnum OtherVerticalPositionTypeEnum { get; set; }
    }
}
