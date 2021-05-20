using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class ItemSearch_DL : Base_DL
    {
        public DataTable PCMN0101K_Search(ItemSearch_Entity ise)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@VC_MAKER_SHOHIN", ise.makercd);
            dic.Add("@VC_KYOGI_CATEGORY", ise.sports);
            dic.Add("@VC_SHOHIN_CATEGORY", ise.category);
            dic.Add("@VC_SHOHIN", ise.itemCode);
            dic.Add("@VC_BRAND", ise.brand);
            dic.Add("@VC_SHIIRESAKI", ise.supplier);
            dic.Add("@SKUCD_S", ise.skucd);
            dic.Add("@vm_hanbai_shohin", ise.itemName);
            dic.Add("@ItemCheck", ise.itemCheck);
            dic.Add("@MakerCheck", ise.makerCheck);
            dic.Add("@VC_JAN1", ise.jancd);
            return SelectData(dic, "PCMN0101K_Search");
        }

        public DataTable PSKS0110S_Select(ItemSearch_Entity ise)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ItemName", ise.itemName);
            dic.Add("@SpecialFlg", ise.specialFlag);
            dic.Add("@YoyakuFlg", ise.yoyakuFlg);
            dic.Add("@SportsCD", ise.sports);
            dic.Add("@BunruiCD", ise.bunrui);
            dic.Add("@SupplierCD", ise.supplier);
            dic.Add("@BarndCD", ise.brand);
            dic.Add("@Nendo", ise.nendo);
            dic.Add("@Season", ise.season);
            dic.Add("@Catalog", ise.catalog);
            dic.Add("@Shijisho", ise.shijisho);
            dic.Add("@MakerCD", ise.makercd);
            dic.Add("@ItemCD", ise.itemCode);
            dic.Add("@JANCD", ise.jancd);
            dic.Add("@WebTenpo", ise.webTenpo);
            dic.Add("@zaikoExists", ise.zaikoExists);
            dic.Add("@SKUCD", ise.skucd);
            dic.Add("@ItemCheck", ise.itemCheck);
            dic.Add("@MakerCheck", ise.makerCheck);
            dic.Add("@MakerLikeSearch", ise.makerLikeSearch);

            return SelectData(dic, "PSKS0110S_Select");
        }

        public DataTable PSKS0110S_Search_Select(ItemSearch_Entity ise)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ItemName", ise.itemName);
            dic.Add("@SpecialFlg", ise.specialFlag);
            dic.Add("@YoyakuFlg", ise.yoyakuFlg);
            dic.Add("@SportsCD", ise.sports);
            dic.Add("@BunruiCD", ise.bunrui);
            dic.Add("@SupplierCD", ise.supplier);
            dic.Add("@BarndCD", ise.brand);
            dic.Add("@Nendo", ise.nendo);
            dic.Add("@Season", ise.season);
            dic.Add("@Catalog", ise.catalog);
            dic.Add("@Shijisho", ise.shijisho);
            dic.Add("@MakerCD", ise.makercd);
            dic.Add("@ItemCD", ise.itemCode);
            dic.Add("@JANCD", ise.jancd);
            dic.Add("@MakerLikeSearch", ise.makerLikeSearch);

            return SelectData(dic,"PSKS0110S_Search_Select");
        }

        public DataTable PSKS0104I_Select(ItemSearch_Entity ise)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@VC_SHIIRESAKI", ise.makercd);
            dic.Add("@InportDateTime", ise.InportDate);
            dic.Add("@VC_BRAND", ise.brand);
            dic.Add("@VO_CATALOG", ise.catalog);
            dic.Add("@VK_NENDO", ise.nendo);
            dic.Add("@VK_SEASON", ise.season);
            dic.Add("@Item", ise.itemCode);
            dic.Add("@SKUCD", ise.skucd);
            dic.Add("@VC_JAN1", ise.jancd);
            dic.Add("@ItemCheck", ise.itemCheck);
            dic.Add("@MakerCheck", ise.makerCheck);
            dic.Add("@Zaiko0", ise.Zaiko0);
            dic.Add("@FirstInport", ise.FirstInport);
            dic.Add("@VC_MAKER_SHOHIN", ise.makerShohin);

            return SelectData(dic, "PSKS0104I_SELECT");
        }

        public DataTable PSKS0120S_Select(ItemSearch_Entity ise,string OperatorCD)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", OperatorCD);
            dic.Add("@ItemName", ise.itemName);
            dic.Add("@SpecialFlg", ise.specialFlag);
            dic.Add("@YoyakuFlg", ise.yoyakuFlg);
            dic.Add("@SportsCD", ise.sports);
            dic.Add("@BunruiCD", ise.bunrui);
            dic.Add("@SupplierCD", ise.supplier);
            dic.Add("@BarndCD", ise.brand);
            dic.Add("@Nendo", ise.nendo);
            dic.Add("@Season", ise.season);
            dic.Add("@Catalog", ise.catalog);
            dic.Add("@Shijisho", ise.shijisho);
            dic.Add("@MakerCD", ise.makercd);
            dic.Add("@ItemCD", ise.itemCode);
            dic.Add("@JANCD", ise.jancd);
            dic.Add("@MakerLikeSearch", ise.makerLikeSearch);

            return SelectData(dic, "PSKS0120S_Select");
        }
    }
}
