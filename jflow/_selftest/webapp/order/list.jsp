<%@ page contentType="text/html; charset=UTF-8" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<html>
<body>
<form action="${ctx}/order/save" method="post">
    <input name="amount"/>
    <button type="submit">save</button>
</form>
<a href="${ctx}/order/list">refresh</a>
</body>
</html>
