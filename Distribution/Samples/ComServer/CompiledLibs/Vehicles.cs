using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace SimpleComServer
{
    [Guid("068E07F7-8D70-4681-83B3-8867136829E7")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class VehicleFactory
    {
        public Car MakeCar(string name)
        {
            Car c = new Car();
            c.Name = name;
            return c;
        }

        public Bicycle MakeBicycle()
        {
            return new Bicycle();
        }
    }

    public abstract class Vehicle
    {
        public abstract int HowManyWheels
        {
            get;
        }
        public abstract string GetSound();
    }

    public class Bicycle : Vehicle
    {
        public override int  HowManyWheels
        {
	        get 
	        { 
		         return 2;
	        }
        }

        public override string GetSound()
        {
            return "Tring!";
        }
    }

    public class Car : Vehicle
    {
        private string _name;
        public string Name { get { return _name; } set { _name = value;} }
        public override int HowManyWheels
        {
            get
            {
                return 4;
            }
        }

        public override string GetSound()
        {
            return "Beep!";
        }
    }
}
