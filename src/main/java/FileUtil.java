import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class FileUtil {

    /**
     * 读取目录下的所有文件
     *
     * @param dir
     *            目录
     * @param fileNames
     *            保存文件名的集合
     * @return
     */
    public static void findFileList(File dir, List<String> fileNames) {
        if (!dir.exists() || !dir.isDirectory()) {// 判断是否存在目录
            return;
        }
        String[] files = dir.list();// 读取目录下的所有目录文件信息
        for (int i = 0; i < files.length; i++) {// 循环，添加文件名或回调自身
            File file = new File(dir, files[i]);
            if (file.isFile()) {// 如果文件
                fileNames.add(dir + "\\" + file.getName());// 添加文件全路径名
            } else {// 如果是目录
                findFileList(file, fileNames);// 回调自身继续查询
            }
        }
    }


    public static List<String>  getList(String patha){
        String path=patha;
        File file=new File(path);
        File[] tempList = file.listFiles();
        System.out.println("该目录下对象个数："+tempList.length);
        List<Map<String, String>> list = new ArrayList<Map<String,String>>();
        List<String> lists = new ArrayList<String>();
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile()) {

                System.out.println("文     件："+tempList[i]);
            }
            if (tempList[i].isDirectory()) {
                lists.add(tempList[i].getPath());
            }
        }
        return lists;
    }
}
