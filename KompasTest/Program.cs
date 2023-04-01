using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;
using KompasAPI7;
using System.Drawing;
using System.Runtime.InteropServices;

namespace KompasTest
{
    internal class Program
    {
       
        static void Main(string[] args)
        {
            Ko ko = new Ko();
            bool exit = false;
            string flag;
            while (!exit)
            {
                //  ko.UseEntityCollection();
                // ko.GetSetPartName();
                //   ko.GetSetUserParamComponent();
                ko.GetAssemblyAndPartsOLD();

                Console.WriteLine("Для выхода нажать 1");
                flag=Console.ReadLine();
                if(flag=="1") exit=true;
            }
        }

        public class Ko
        {
            private KompasObject kompas;
            private ksDocument3D doc;
            private string buf;
            // Взять/изменить имя компоненты
          public  void GetSetPartName()
            {
                double mass;
                string material;

                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D(); // привязываемся к активному документу

                ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);  // верхний компонент
                if (part != null)
                {
                    buf = string.Format("Имя компоненты {0}", part.name);
                    kompas.ksMessage(buf);
                    part.name = "Втулка";

                    mass = Math.Round( part.GetMass(),3);
                    material = part.material;
                  

                    part.Update();

                    Console.WriteLine($"{mass} {material}");
                }
            }

            // Взять и поменять внешние переменные компоненты
            void GetSetArrayVariable()
            {
                if (doc.IsDetail())
                {
                    kompas.ksError("Текущий документ должен быть сборкой");
                    return;
                }
                ksPart part = (ksPart)doc.GetPart(0);   // первая деталь в сборке
                if (part != null)
                {
                    // работа с массивом внешних переменных
                    ksVariableCollection varCol = (ksVariableCollection)part.VariableCollection();
                    if (varCol != null)
                    {
                        ksVariable var = (ksVariable)kompas.GetParamStruct((short)StructType2DEnum.ko_VariableParam);
                        if (var == null)
                            return;
                        for (int i = 0; i < varCol.GetCount(); i++)
                        {
                            var = (ksVariable)varCol.GetByIndex(i);
                            buf = string.Format("Номер переменной {0}\nИмя переменной {1}\nЗначение переменной {2}\nКомментарий {3}", i, var.name, var.value, var.note);
                            kompas.ksMessage(buf);
                            if (i == 0)
                            {
                                var.note = "qwerty";
                                double d = 0;
                                kompas.ksReadDouble("Введи переменную", 10, 0, 100, ref d);
                                var.value = d;
                            }
                        }

                        for (int j = 0; j < varCol.GetCount(); j++)
                        {
                            // просмотр изменненных переменных
                            var = (ksVariable)varCol.GetByIndex(j);
                            buf = string.Format("Номер переменной {0}\nИмя переменной {1}\nЗначение переменной {2}\nКомментарий {3}", j, var.name, var.value, var.note);
                            kompas.ksMessage(buf);
                        }
                        part.RebuildModel();    // перестроение модели
                    }
                }
            }
            // Установить и получить параметры пользователя в компоненте
           public void GetSetUserParamComponent()
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D(); // привязываемся к активному документу

                if (doc.IsDetail())
                {
                    kompas.ksError("Текущий документ должен быть сборкой");
                    return;
                }

                ksPart part = (ksPart)doc.GetPart(0);   // первая деталь в сборке
                ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
                if (par == null || item == null || arr == null || part == null)
                    return;

                par.Init();
                par.SetUserArray(arr);
                item.Init();
                item.doubleVal = 12.12;
                arr.ksAddArrayItem(-1, item);
                item.Init();
                item.doubleVal = 21.21;
                arr.ksAddArrayItem(-1, item);
                item.Init();
                item.intVal = 666;
                arr.ksAddArrayItem(-1, item);
                item.Init();
                item.intVal = 999;
                arr.ksAddArrayItem(-1, item);

                part.SetUserParam(par); // установка пользовательской структуры
                part.Update();

                buf = string.Format("Размер пользовательской структуры {0}", part.GetUserParamSize()); // размер пользовательской структуры
                kompas.ksMessage(buf);

                ksUserParam par2 = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                ksLtVariant item2 = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                ksDynamicArray arr2 = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
                if (par2 == null || item2 == null || arr2 == null)
                    return;

                par2.Init();
                par2.SetUserArray(arr2);
                item2.Init();
                item2.doubleVal = 0;
                arr2.ksAddArrayItem(-1, item2);
                item2.Init();
                item2.doubleVal = 0;
                arr2.ksAddArrayItem(-1, item2);
                item2.Init();
                item2.intVal = 0;
                arr2.ksAddArrayItem(-1, item2);
                item2.Init();
                item2.intVal = 0;
                arr2.ksAddArrayItem(-1, item2);

                part.GetUserParam(par2);    // берем пользовательскeую структуру

                dstruct d;

                arr2.ksGetArrayItem(0, item2);
                d.a = item2.doubleVal;
                arr2.ksGetArrayItem(1, item2);
                d.b = item2.doubleVal;
                arr2.ksGetArrayItem(2, item2);
                d.c = item2.intVal;
                arr2.ksGetArrayItem(3, item2);
                d.d = item2.intVal;
                buf = string.Format("Переменные пользовательского масства\na = {0}\nb = {1}\nc = {2}\nd = {3}", d.a, d.b, d.c, d.d);
                kompas.ksMessage(buf);  // просмотрим переменные из пользовательского массива
            }
            // Использование массива элементов
            public void UseEntityCollection()
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D(); // привязываемся к активному документу
                if (doc != null)
                {
                    ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);  // новый компонент
                    if (part != null)
                    {
                        int count1 = 0;
                        int count2 = 0; // количество плоских поверхностей
                        int count = 0;  // количество конических поверхностей
                        double area = 0;
                        double totalArea = 0;

                        // массив поверхностей
                        ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);
                        if (collect != null)
                        {
                            count = collect.GetCount();
                            count1 = 0;
                            count2 = 0;
                            ksColorParam colorPr = null;    // интерфейс свойств цвета
                            if (collect != null && count != 0)
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    ksEntity ent = (ksEntity)collect.GetByIndex(i);
                                    colorPr = (ksColorParam)ent.ColorParam();
                                    // интерфейс свойств поверхности
                                    ksFaceDefinition faceDef = (ksFaceDefinition)ent.GetDefinition();
                                    if (faceDef != null)
                                    {
                                        // коническая по-ть		//цилиндрическая по-ть
                                        if (faceDef.IsCone() || faceDef.IsCylinder())
                                        {
                                            colorPr.color = Color.FromArgb(0, 255, 255, 0).ToArgb();
                                            count2++;   // считаем количество объектов

                                            area = faceDef.GetArea(1);
                                            totalArea += area;
                                        }
                                        // плоская по-ть  
                                        if (faceDef.IsPlanar())
                                        {
                                            colorPr.color = Color.FromArgb(0, 0, 255, 255).ToArgb();
                                            count1++;   // считаем количество объектов

                                            area = faceDef.GetArea(1);
                                            totalArea += area;
                                        }

                                        ent.Update();   // обновить параметры
                                    }
                                }
                            }
                        }

                        // сообщяем о результатах работы
                        if (count == 0)
                            kompas.ksMessage("Не найдено ни одной поверхности");
                        else
                        {
                            totalArea = Math.Round( totalArea * 0.000001,5);
                            buf = string.Format($"Найдено {count2} коничечких и {count1} плоских объектов, Общая площадь {totalArea}" );
                            kompas.ksMessage(buf);

                            Console.WriteLine(totalArea*0.001);

                        }

                        count1 = 0;
                        count2 = 0;
                        // массив ребер
                        ksEntityCollection collect2 = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_edge);
                        count = collect2.GetCount();
                        if (collect2 != null && count != 0)
                        {
                            for (int i = 0; i < count; i++)
                            {
                                ksEntity ent = (ksEntity)collect2.GetByIndex(i);
                                ksEdgeDefinition edgeDef = (ksEdgeDefinition)ent.GetDefinition();
                                if (edgeDef != null)
                                {
                                    if (edgeDef.IsStraight())
                                        count1++;   // количество прямых ребер
                                    else
                                        count2++;   // количество криволинейных ребер
                                }
                            }
                        }

                        // сообщяем о результатах работы
                        if (count == 0)
                            kompas.ksMessage("Не найдено ни одного ребра");
                        else
                        {
                            buf = string.Format("Найдено {0} прямых и {1} криволинейных ребер", count1, count2);
                            kompas.ksMessage(buf);
                        }
                    }
                }
            }
            private struct dstruct
            {
                public double a, b;
                public int c, d;
            }
            //Получить редактировать свойства модели
            public void GetSetProperty()
            {
                IApplication application = (IApplication)Marshal.GetActiveObject("KOMPAS.Application.7");
                IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
                IPart7 part = document3D.TopPart;
                //получаю свойства докумнта
                IPropertyMng propertyMng = (IPropertyMng)application;

                //создаю свойства
                IProperty property = propertyMng.AddProperty(null, null);
                property.Name = "Новое свойство";
                property.Update();

                
                //получаю свойства докумнта
                var properties = propertyMng.GetProperties(document3D);
                foreach (IProperty item in properties)
                {
                    if(item.Name == "Новое свойство")
                    {
                        dynamic info;
                        bool sourse;
                        IPropertyKeeper propertyKeeper = (IPropertyKeeper)item;
                        propertyKeeper.SetPropertyValue((_Property)item, 123, false);
                        item.Update();

                        //Получаю все свойства модели
                        propertyKeeper.GetPropertyValue((_Property)item, out info,false, out sourse);
                        Console.WriteLine(info);
                    }
                }

            }
            //Поход по всем сборкам на aplication 5
            public void GetAssemblyAndPartsOLD()
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D(); // привязываемся к активному документу


                if (doc != null)
                {
                 
                    IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
                    IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;

                    IPropertyMng propertyMng = (IPropertyMng)application;

                    //получаю сборку
                    IPart7 part = document3D.TopPart;
                    
                    IFeature7 f7 = part as IFeature7;
                    if (f7.ResultBodies is object[])
                        foreach (IBody7 body in f7.ResultBodies)
                        {
                            if (body is IBody7)
                            {
                                var properties = propertyMng.GetProperties(document3D);
                                foreach (IProperty item in properties)
                                {
                                    if (item.Name == "Масса")
                                    {
                                        dynamic info;
                                        bool sourse;
                                        IPropertyKeeper propertyKeeper = (IPropertyKeeper)body;

                                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out sourse);

                                        Console.WriteLine(info);
                                    }
                                }
                            }
                        }

                }
            }
          
            //Проход по всем сборкам и подсборкам
            public void GetAssemblyAndParts ()
            {
                IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
                IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
               
                //получаю сборку
                IPart7 part = document3D.TopPart;

                List<IPart7> parts = new List<IPart7> ();

                Recursion recursion = new Recursion ();
                recursion.GetDetails(part,parts);

                foreach (IPart7 item in parts)
                {
                    Console.WriteLine(item.Marking + " - " + item.Name);
                }
            }
            /// <summary>
            /// Класс для рекурсивного метода создания списка деталей всех сборок
            /// </summary>
            class Recursion
            {
                public void GetDetails (IPart7 part, List<IPart7>parts)
                {

                    parts.Add(part);
                    foreach(IPart7 item in part.Parts)
                    {
                        if (item.Detail == true) parts.Add(item);
                        if (item.Detail == false) GetDetails(item, parts);

                     //   ksBodyParts body=item.GetBodyById(0);
                       // ksEntity body=(ksEntity)item.getd
                    }
                }
            }
        }
    }
}
