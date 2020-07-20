import java.io.File
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import scala.collection.JavaConverters._

object HarvestReport{
  def main(args: Array[String]): Unit ={
    //println("To check the Harvest report for us and other countries")
    val empFile = "src\\main\\resources\\InternationalBaseline2019-Final.xlsx"

    val fis = new FileInputStream(myFile)
    val myWorkbook = new HSSFWorkbook(fis)
    val mySheet = myWorkbook.getSheetAt(0)
    val rowIterator = mySheet.iterator()
    while(rowIterator.hasNext){
      val row = rowIterator.next()
      val cellIterator = row.cellIterator()
      while(cellIterator.hasNext) {
        val cell = cellIterator.next()
        cell.getCellType match {
          case Cell.CELL_TYPE_STRING => {
            print(cell.getStringCellValue + "\t")
          }
          case Cell.CELL_TYPE_NUMERIC => {
            print(cell.getNumericCellValue + "\t")
          }
          case Cell.CELL_TYPE_BOOLEAN => {
            print(cell.getBooleanCellValue + "\t")
          }
          case Cell.CELL_TYPE_BLANK => {
            print("null" + "\t")
          }
          case _ => throw new RuntimeException(" this error occured when reading ")
          //        case Cell.CELL_TYPE_FORMULA => {print(cell.getF + "\t")}
        }
      }
      println("")
    }
      }
  }
