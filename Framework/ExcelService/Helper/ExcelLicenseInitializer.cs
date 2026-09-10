using OfficeOpenXml;
namespace ExcelService.Helper
{
    public static class ExcelLicenseInitializer
    {
        public static void Initialize()
        {
            ExcelPackage.License.SetNonCommercialOrganization(AppConstant.EXCEL_HELPER_LICENSING_ORGANIZATION);
        }
    }
}
