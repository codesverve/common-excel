package com.uetty.common.excel.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 反射工具类
 * @author vince
 */
public class ReflectUtil {
	
	private static Logger logger = LoggerFactory.getLogger(ReflectUtil.class);

	private static String getterName(String fieldName) {
		return "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
	}
	
	private static String setterName(String fieldName) {
		return "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
	}

	private static String getNameFromSetterName(String setterName) {
        String name = setterName.substring(3);
        return name.substring(0,1).toLowerCase() + name.substring(1);
    }
	
	public static Object getFieldValue(Object obj, String fieldName) {
		Class<? extends Object> clz = obj.getClass();
		try {
			String getterName = getterName(fieldName);
			return invokeMethod(obj, getterName);
		} catch (Exception e) {}
		try {
			Field field = clz.getDeclaredField(fieldName);
			if (field != null) {
				field.setAccessible(true);
				return field.get(obj);
			}
		} catch (Exception e) {};
		try {
			Field field = clz.getField(fieldName);
			if (field != null) {
				return field.get(obj);
			}
		} catch (Exception e) {}
		throw new RuntimeException("field[" + fieldName + "] not found in " + obj);
	}
	
	public static void setFieldValue(Object obj, String fieldName, Object value) {
		Class<? extends Object> clz = obj.getClass();

		try {
			String setterName = setterName(fieldName);
			invokeMethod(obj, setterName, value);
			return;
		} catch (Exception e) {}
		
		try {
			Field field = clz.getDeclaredField(fieldName);
			if (field != null) {
				field.setAccessible(true);
				field.set(obj, value);
				return;
			}
		} catch (Exception e) {}
		try {
			Field field = clz.getField(fieldName);
			if (field != null) {
				field.set(obj, value);
				return;
			}
		} catch (Exception e) {}
		throw new RuntimeException("field[" + fieldName + "] not found in " + obj);
	}
	
	public static List<String> getFieldNames(Object obj) {
		Class<? extends Object> clz = obj.getClass();
		Field[] fields = clz.getDeclaredFields();
		List<String> list = new ArrayList<String>();
		for (Field field : fields) {
			String name = field.getName();
			list.add(name);
		}
		return list;
	}

//
//	public static List<String> getFieldNames(Class<?> clz) {
//		Field[] fields = clz.getDeclaredFields();
//		Set<String> set = Arrays.stream(clz.getDeclaredFields()).map(Field::getPropValue1).collect(Collectors.toSet());
//		Set<String> collect = Arrays.stream(clz.getFields()).map(Field::getPropValue1).collect(Collectors.toSet());
//		set.addAll(collect);
//		return new ArrayList<>(set);
//	}

	public static List<Field> getFields(Class<?> clz) {
		Field[] declaredFieldArr = clz.getDeclaredFields();
		List<Field> fields = Arrays.asList(declaredFieldArr);

		Field[] fieldArr = clz.getFields();
        List<Field> fields1 = Arrays.asList(fieldArr);

		fields.addAll(fields1);
		return fields.stream().distinct().collect(Collectors.toList());
	}

	public static Field getFieldByName(Class<?> clz, String name) {
	    Field field = null;
	    try {
            field = clz.getDeclaredField(name);
        } catch (Exception ignore) {}
	    if (field != null) return field;
	    try {
	        field = clz.getField(name);
        } catch (Exception ignore) {}
	    if (field != null) return field;

        Class<?> superclass = clz.getSuperclass();
        if (superclass == null || superclass.equals(Object.class)) return null;
        return getFieldByName(superclass, name);
    }

	public static List<Field> getFieldsBySetter(Class<?> clz) {
        Method[] methodArr = clz.getMethods();
        return Arrays.stream(methodArr).filter(m -> m.getName().startsWith("set")).map(m -> {
            String name = getNameFromSetterName(m.getName());
            return getFieldByName(clz, name);
        }).distinct().collect(Collectors.toList());
    }

	public static List<Field> getDeclaredFields(Class<?> clz) {
		List<Field> list = new ArrayList<>();
		for(Class tempClass = clz; tempClass != null; tempClass = tempClass.getSuperclass()) {
			list.addAll(Arrays.asList(tempClass.getDeclaredFields()));
		}
		return list;
	}

	public static Class<?> getFieldClass(Object obj, String fieldName) {
		Class<? extends Object> clz = obj.getClass();
		try {
			Field field = clz.getDeclaredField(fieldName);
			if (field != null) {
				field.setAccessible(true);
				return field.getType();
			}
		} catch (Exception e) {};
		try {
			Field field = clz.getField(fieldName);
			if (field != null) {
				return field.getType();
			}
		} catch (Exception e) {}
		throw new RuntimeException("field[" + fieldName + "] not found in " + obj);
	}

	public static Object getInstance(Class<? extends Object> clz, Object... params) {
		Constructor<?>[] constructors = clz.getConstructors();
		for (Constructor<?> constructor : constructors) {
			try {
				Class<?>[] parameterTypes = constructor.getParameterTypes();
				if (parameterTypes.length != params.length) {
					continue;
				}
				constructor.setAccessible(true);
				return constructor.newInstance(params);
			} catch (Exception e) {}
		}
		if (params.length == 0) {
			try {
				return clz.newInstance();
			} catch (Exception e) {}
		}

		return throwMethodNotFound(clz, 1, params);
	}

	public static Object invokeMethod(Object obj, String methodName, Object... params) {
		Class<? extends Object> clz = obj.getClass();
		if (params.length == 0) {
			try {
				Method method = clz.getMethod(methodName);
				return method.invoke(obj);
			} catch (Exception e) {}
		} else {
			// 参数类型绝对匹配的方法查找
			Class<?> pclzs[] = new Class<?>[params.length];
			boolean noNull = true;
			for (int i = 0; i < params.length; i++) {
				Object p = params[i];
				if (p == null) {
					noNull = false;
				} else {
					pclzs[i] = p.getClass();
				}
			}
			if (noNull) {
				try {
					Method method = clz.getMethod(methodName, pclzs); // 参数类型绝对匹配的方法
					return method.invoke(obj, params);
				} catch (Exception e) {}
			}
			
			// 不追求绝对匹配，只要能调用即可
			Method[] methods = clz.getMethods();
			for (Method m : methods) {
				try {
					if (!m.getName().equals(methodName)) continue;
					Class<?>[] parameterTypes = m.getParameterTypes();
					if (parameterTypes.length != params.length) continue;
					return m.invoke(obj, params);
				} catch (Exception e) {}
			}
		}
		
		return throwMethodNotFound(clz, 3, params);
	}

	private static Object throwMethodNotFound(Class<?> clz, int methodType, Object[] params) {
		String errorMsg = "(";
		if (methodType == 1) {
			errorMsg = "constructor " + errorMsg;
		} else if (methodType == 2) {
			errorMsg = "static method " + errorMsg;
		} else if (methodType == 3) {
			errorMsg = "method " + errorMsg;
		}
		for (int i = 0; i < params.length; i++) {
			if (i != 0) errorMsg += ", ";
			errorMsg += params[i] == null ? "null" : params[i].getClass().getName();
		}
		errorMsg += ") not found in class[" + clz.getCanonicalName() + "]";
		throw new RuntimeException(errorMsg);
	}

	public static Object invokeMethod(Class<?> clz, String methodName, Object... params) {
		if (params.length == 0) {
			try {
				Method method = clz.getMethod(methodName);
				return method.invoke(null);
			} catch (Exception e) {}
		} else {
			// 参数类型绝对匹配的方法查找
			boolean noNull = true;
			Class<?> pclzs[] = new Class<?>[params.length];
			for (int i = 0; i < params.length; i++) {
				Object p = params[i];
				if (p == null) {
					noNull = false;
					break;
				}
				pclzs[i] = p.getClass();
			}
			if (noNull) {
				try {
					Method method = clz.getMethod(methodName, pclzs); // 参数类型绝对匹配的方法
					return method.invoke(null, params);
				} catch (Exception e) {}
			}
			
			// 不追求绝对匹配，只要能调用即可
			Method[] methods = clz.getMethods();
			for (Method m : methods) {
				try {
					if (!m.getName().equals(methodName)) continue;
					Class<?>[] parameterTypes = m.getParameterTypes();
					if (parameterTypes.length != params.length) continue;
					return m.invoke(null, params);
				} catch (Exception e) {}
			}
		}
		
		return throwMethodNotFound(clz, 2, params);
	}
	
	/**
	 * 打印该类包含的变量名
	 */
	public static void printContainFieldNames(Class<?> clz) {
		Set<String> fieldSet = new HashSet<String>();
		try {
			Field[] fields = clz.getFields(); // 公共变量（包含自父类继承的变量）
			for (Field field : fields) {
				fieldSet.add(field.getName());
			}
		} catch (Throwable e) {
			logger.error(e.getMessage(), e);
		}
		try {
			Field[] fields = clz.getDeclaredFields(); // 非公共变量（无法包含自父类继承的变量）
			for (Field field : fields) {
				fieldSet.add(field.getName());
			}
		} catch (Throwable e) {
			logger.error(e.getMessage(), e);
		}
		for (String string : fieldSet) {
			logger.debug(string);
		}
	}
	
	/**
	 * 打印类所在的文件路径
	 * <p> 适合代码多个有包含相同类的jar出现bug时，借用该方法排除
	 */
	public static void printClassPath(Class<?> clz) {
		String classFilePath = clz.getName();
		classFilePath = classFilePath.replace('.', '/');
		classFilePath += ".class";
		
		printClassLoaderAndPath(clz.getClassLoader(), clz, classFilePath);
	}
	
	private static boolean hasClass(URL url, String classFilePath) {
		try {
			URL[] uls = {url};
			@SuppressWarnings("resource")
			URLClassLoader myLoader = new URLClassLoader(uls, null);
			URL resource = myLoader.getResource(classFilePath);
			return resource != null;
		} catch (Exception e) {
			return false;
		}
	}
	
	private static void printClassLoaderAndPath(ClassLoader classloader, Class<?> clz, String classFilePath) {
		if (classloader == null) {
			return;
		}
		logger.debug("classloader ==> " + classloader.toString());
		
		if (!(classloader instanceof URLClassLoader)) {
			return;
		}
		
		URLClassLoader urlClassLoader = (URLClassLoader) classloader;
		URL[] urls = urlClassLoader.getURLs();
		
		for (URL url : urls) {
			if (!hasClass(url, classFilePath)) {
				continue;
			}
			logger.debug(url.getPath());
		}
		logger.debug("------------------------------------------------");
		
		ClassLoader parent = classloader.getParent();
		printClassLoaderAndPath(parent, clz, classFilePath);
	}

}