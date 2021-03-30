using InfrastructureLibary.DataBase;

namespace InfrastructureLibary.Constants
{
    public static class Global
    {
        private static efDbContext _dataBase;

        public static efDbContext DataBase { get => GetDataBase(); set => _dataBase = value; }


        private static efDbContext GetDataBase()
        {
            if (_dataBase == null)
            {
                _dataBase = new efDbContext();
                _dataBase.Database.EnsureCreated();
            }
            return _dataBase;
        }
    }
}
