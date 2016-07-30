<form action="../index.php" method="post" enctype="multipart/form-data">
    <table>
        <tbody>
        <tr>
            <th>操作</th>
            <td>
                <select name="action">
                    <option value="read">读取</option>
                    <option value="write">写入</option>
                </select>
            </td>
        </tr>
        <tr>
            <th>读取文件/模板</th>
            <td><input type="file" name="template" /></td>
        </tr>
        <tr>
            <th>xlscript 脚本</th>
            <td><textarea name="xlscript"></textarea></td>
        </tr>
        <tr>
            <td colspan="2"><input type="submit" /></td>
        </tr>
        </tbody>
    </table>
</form>