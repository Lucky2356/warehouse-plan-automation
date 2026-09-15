using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class ReceivingFormulasTests
{
    [Fact]
    public void ВПриемку_КакВШаблонеАналитика()
    {
        // «итог»: «Код» в G, «Допоставить» в S; на «А2, А3» «Код» в D, данные со второй строки.
        Assert.Equal("=SUMIF(итог!G:G,D2,итог!S:S)", ReceivingFormulas.ToReceive("итог", 7, 19, 4, 2));
        Assert.Equal("=СУММЕСЛИ(итог!G:G;D2;итог!S:S)", ReceivingFormulas.ToReceiveLocal("итог", 7, 19, 4, 2));
    }

    [Fact]
    public void КолонкиБерутсяИзЗаголовков_ЛистСПробеломВКавычках()
    {
        Assert.Equal(
            "=SUMIF('итог 10.09'!H:H,E3,'итог 10.09'!T:T)",
            ReceivingFormulas.ToReceive("итог 10.09", 8, 20, 5, 3));
    }
}
