from enum import Enum
from functools import cached_property
from typing import Any, Dict, Optional, Union, List, TYPE_CHECKING
from invlog_utils.common.database_common import DBMS, DatabaseConnection
from pyspark.sql import SparkSession, DataFrame, DataFrameReader, DataFrameWriter
import logging

if TYPE_CHECKING:
    from invlog_utils.common.database_common import ConnectionConfig
    
class SaveMode(Enum):
    APPEND = 0
    """append contents of this :class:Dataframe to existing data."""
    OVERWRITE = 1
    """overwrite existing data."""
    IGNORE = 2
    """Silently ignore this operation if data already exists."""
    ERROR = 3
    """Throw an exception if data already exists."""

class PartitionSchema:
    """Defines a policy for splitting database table data into Spark partitions."""
    def __init__(self,
                 num_partitions: int = 1,
                 partition_column: str = None,
                 lower_bound: Union[int, float] = None,
                 upper_bound: Union[int, float] = None,
                 predicates: Optional[list] = None):
        self.num_partitions = num_partitions
        self.partition_column = partition_column
        self.lower_bound = lower_bound
        self.upper_bound = upper_bound
        self.predicates = predicates
        
    @staticmethod
    def builder() -> "PartitionSchema":
        return PartitionSchema()
    
    def set_num_partitions(self, numPartitions: int) -> "PartitionSchema":
        self.num_partitions = numPartitions
        return self
    
    def set_partition_column(self, partitionColumn: str) -> "PartitionSchema":
        """The column that will be used to partition the data across Spark Workers

        Args:
            partitionColumn (str): The column that will be used to partition the 
            data across Spark workers. Supported column types include numeric, date,
            and timestamp.

        Returns:
            PartitionSchema: self
        """
        self.partition_column = partitionColumn
        return self
    
    def set_lower_bound(self, lowerBound: Union[int, float]) -> "PartitionSchema":
        self.lower_bound = lowerBound
        return self
    
    def set_upper_bound(self, upperBound: Union[int, float]) -> "PartitionSchema":
        self.upper_bound = upperBound
        return self
    
    def set_predicates(self, predicates: list) -> "PartitionSchema":
        self.predicates = predicates
        return self
    
    def _is_predicates_based(self) -> bool:
        return self.predicates is not None and len(self.predicates) > 0
    
    def _validate(self):
        if self._is_predicates_based():
            self._validate_predicates()
        else:
            self._validate_range()
            
    def _validate_predicates(self):
        if not isinstance(self.predicates, list):
            raise Exception('predicates must be a list of SQL strings.')
        if len(self.predicates) == 0:
            raise Exception('predicates list cannot be empty.')
        for predicate in self.predicates:
            if not isinstance(predicate, str) or len(predicate) == 0:
                raise Exception(f'Each predicate must be a non-empty string, got: {predicate}')
    
    def _validate_range(self):
        if self.num_partitions <= 0:
            raise Exception(f'num_partitions must be a >= 0 but was {self.num_partitions}')
        if self.partition_column is None or len(self.partition_column) == 0:
            raise Exception('partition_column must be a valid column.')
        if self.lower_bound is None:
            raise Exception('lower_bound must be populated.')
        if self.upper_bound is None:
            raise Exception('upper_bound must be populated.')
        if self.lower_bound >= self.upper_bound:
            raise Exception(f'lower_bound must be less than upper_bound, got: {self.lower_bound} >= {self.upper_bound}')
        
    def _get_partition_options(self) -> dict:
        return {
            "column": self.partition_column,
            "lowerBound": self.lower_bound,
            "upperBound": self.upper_bound,
            "numPartitions": self.num_partitions
        }
    
    def _get_predicates(self) -> List[str]:
        self._validate_predicates()
        return self.predicates

class SparkDatabaseConnectionInterface(DatabaseConnection):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.spark = None
    
    @property
    def iceberg(self) ->  'SparkDatabaseConnectionIceberg':
        """Cast this connection to :class:`SparkDatabaseConnectionIceberg`.
        
        Raises:
            TypeError: If this connection is not an Iceberg connection. Check that 
            the corresponding ``[Databases]`` entry has ``dbms = ICEBERG``.
        """
        raise TypeError(
            f"{self.__class__.__name__} is not an Iceberg connection. " 
            "Ensure the correspondng [Database] entry has dbms = ICEBERG."
        )
        
    @classmethod
    def read_table_into_spark(
        self,
        table:str,
        partition:PartitionSchema=None,
        fetch_size:Optional[int]=None
    ) -> DataFrame:
        pass

    @classmethod
    def write_table_from_spark(
        self,
        df:DataFrame,
        table:str,
        mode:SaveMode=SaveMode.ERROR,
        max_partitions:Optional[int]=None,
        batch_size:Optional[int]=None,
        truncate=False        
    ):
        pass
    
class SparkDatabaseConnection(SparkDatabaseConnectionInterface):
    """Stores the connection information for a Database and provides some functions for connecting and interacting with the database.
    """
    def __init__(self,
                 username:str,
                 password:str,
                 host:str,
                 port:int,
                 database:str,
                 dbms:DBMS,
                 oracle_lib_dir:Optional[str]=None):
        super().__init__(username, password, host, port, database, dbms)
        self.oracle_lib_dir = oracle_lib_dir
        match dbms:
            case DBMS.ORACLE:
                self.jdbc_url = f"jdbc:oracle:thin:@(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL=TCP)(HOST={self.host})(PORT={self.port})))(CONNECT_DATA=(SERVICE_NAME={self.database})(SERVER=DEDICATED)))"
                self.jdbc_driver = "oracle.jdbc.OracleDriver"
            case DBMS.SQREAM:
                self.jdbc_url = f"jdbc:Sqream://{self.host}:{self.port}/{self.database};cluster=true"
                self.jdbc_driver = "com.sqream.jdbc.SQDriver"
            case DBMS.DENODO:
                self.jdbc_url = f"jdbc:denodo://{self.host}:{self.port}/{'common' if self.database is None else self.database}"
                self.jdbc_driver = "com.denodo.vdp.jdbc.Driver"            
            case DBMS.POSTGRES:
                self.jdbc_url = f"jdbc:postgresql://{self.host}:{self.port}/{self.database}"
                self.jdbc_driver = "org.postgresql.Driver"
            case _:
                raise ValueError(f"Unsupported DBMS: {dbms}")

    @classmethod
    def from_config(cls,
                    name:str,
                    config:'ConnectionConfig',
                    oracle_lib_dir:Optional[str]=None) -> SparkDatabaseConnection:
        """Create SparkDatabaseConnection from parsed config object.

        Args:
            name (str): Connection name (used for identificaiton).
            config (ConnectionConfig): ConnectionConfig object containing database conneciton parameters.
            oracle_lib_dir: Optional Oracle client library directory.

        Returns:
            Fully initialized SparkDatabaseConnection.
            
        Raises:
            ValueError: If the config.dbms is ICEBERG.
        """

        if config.dbms == DBMS.ICEBERG:
            raise ValueError("Cannot create SparkDatabaseConnection for ICEBERG DBMS. Use SparkDatabaseConnectionIceberg instead.")     
        return cls(
            username=config.username,
            password=config.password,
            host=config.host,
            port=config.port,
            database=config.database,
            dbms=config.dbms,
            oracle_lib_dir=oracle_lib_dir
        )       

    def read_table_into_spark(
        self,
        table:str,
        partition:PartitionSchema=None,
        fetch_size:Optional[int]=None
    ) -> DataFrame:
        return self.query_into_spark(table, partition, fetch_size)

    def query_into_spark(
        self,
        query:str,
        partition:PartitionSchema=None,
        fetch_size:Optional[int]=None,
        **kwargs
    ) -> DataFrame:
        """Execute and load a query against this database connection into a Spark DataFrame.

        Args:
            query (str): The SQL query string to execute.
            partition (PartitionSchema, optional): Partitioning schema for the query. Defaults to None.
            fetch_size (int, optional): Number of rows to fetch per round trip. Defaults to None.
            **kwargs: Additional keyword arguments for Spark JDBC read.

        Returns:
            DataFrame: A Spark DataFrame containing the results of the query.
        """
        
        properties = {
            "user": self.username,
            "password": self.password,
            "driver": self.jdbc_driver
        }
    
        if fetch_size is not None:
            properties["fetchsize"] = str(fetch_size)
            
        for key, value in kwargs.items():
            properties[key] = str(value)
                
        if partition is not None and partition._is_predicates_based():
            print("Predicate based partitioning")
            return self.spark.read.jdbc(
                url=self.jdbc_url,
                table=query,
                predicates=partition._get_predicates(),
                properties=properties
            )
        else:
            if partition is not None:
                partition_options = partition._get_partition_options()
            else:
                partition_options = {}


            return self.spark.read.jdbc(
                url=self.jdbc_url,
                table=query,
                properties=properties,
                **partition_options
            )
            
            
    def write_table_from_spark(
        self,
        df:DataFrame,
        table:str,
        mode:SaveMode=SaveMode.ERROR,
        batch_size:Optional[int]=None,
        numPartitions:Optional[int]=None,
        truncate=False,
        **kwargs:Dict[str, Any]
    ):
        """Writes a Spark Dataframe to a database table using the connection represented by this object.

        Args:
            df (DataFrame): The Spark DataFrame to write.
            table (str): The name of the table to load data into.
            mode (SaveMode, optional): The protocol to use for loading the data. Defaults to SaveMode.ERROR.
            batch_size (Optional[int], optional): The batch size per insert. Defaults to None.
            truncate (bool, optional):
            
            Additional Spark supported JDBC write arguments are also supported. The complete list of args can be found here: 
            https://spark.apache.org/docs/latest/sql-data-sources-jdbc.html
        """
        
        write:DataFrameWriter = (
            df.write.format("jdbc")
               .option("url", self.jdbc_url)
               .option("dbtable", table)
               .option("user", self.username)
               .option("password", self.password)
               .option("driver", self.jdbc_driver)
               .mode(mode.name.lower())
        )
        if truncate is not None:
            write = write.option("truncate", truncate)
        if batch_size is not None:
            write = write.option("batchsize", batch_size)
        if numPartitions is not None:
            write = write.option("numPartitions", numPartitions)
        
        for key, value in kwargs:
            write = write.option(key, value)
        
        write.save()
        
    @cached_property
    def sql(self):
        from invlog_utils.common.database import SQLAlchemyDatabaseConnection
        sql = SQLAlchemyDatabaseConnection(
            username=self.username,
            password=self.password,
            host=self.host,
            port=self.port,
            database=self.database,
            dbms=self.dbms,
            oracle_lib_dir=self.oracle_lib_dir
        )
        return sql
    

class SparkDatabaseConnectionIceberg(SparkDatabaseConnectionInterface):
    """Stores connection information for an Apache Iceberg catalog and provides read/write
    helpers that use Spark's natice Iceberg data source.
    """
    
    def __init__(
        self,
        username:str,
        password:str,
        host:str,
        port:int,
        database:str,
        catalog_name:str,
        catalog_type:str="hadoop",
        catalog_uri:Optional[str]=None,
        warehouse_path:Optional[str]=None,
    ):
        super().__init__(username, password, host, port, database, DBMS.ICEBERG)
        self.catalog_name = catalog_name
        self.catalog_type = catalog_type.lower()
        self.catalog_uri = catalog_uri
        self.warehouse_path = warehouse_path or "/user/hive/warehouse"
        
    @classmethod
    def from_config(cls, name:str, config:'ConnectionConfig') -> SparkDatabaseConnectionIceberg:
        """Create SparkDatabaseConnectionIceberg from parsed config object.

        Args:
            name (str): Connection name (used as catalog name).
            config (ConnectionConfig): ConnectionConfig object containing database conneciton parameters and Iceberg settings.
        
        Returns:
            Fully initialized SparkDatabaseConnectionIceberg.
        
        Raises:
            ValueError: If the config.dbms is not ICEBERG.
        """
        if config.dbms != DBMS.ICEBERG:
            raise ValueError(f"Expected DBMS.ICEBERG, got {config.dbms}.")
        return cls(
            username=config.username,
            password=config.password,
            host=config.host,
            port=config.port,
            database=config.database,
            catalog_name=name,
            catalog_type=config.catalog_type or "hadoop",
            catalog_uri=config.catalog_uri,
            warehouse_path=config.warehouse_path or '/user/hive/warehouse'
        )
        
    @property
    def iceberg(self) ->  'SparkDatabaseConnectionIceberg':
        return self
    
    def spark_builder_configs(self) -> dict:
        """Returns the Spark config key/value pairs required for this Iceberg catalog.
        
        These must be applied to ``SparkSession.builder`` before ``getOrCreate()`` is
        called - catalog extensions and catalog implementations are startup-time settings
        that Spark ignores if set after the session is created.
        
        ``SparkJob`` picks these up automatically during session construction
        
        The catalog is registered under ``self.catalog_name`` (derived from the connnection
        name in ``[Database]``), so multiple Iceberg connections each get their own independently
        -registered Spark catalog with no key collisions.
        """
        configs = {
            "spark.sql.extensions": "org.apache.iceberg.spark.extensions.IcebergSparkSessionExtensions",
            f"spark.sql.catalog.{self.catalog_name}": "org.apache.iceberg.spark.SparkCatalog",
            f"spark.sql.catalog.{self.catalog_name}.type": self.catalog_type
        }
        if self.catalog_type == "rest":
            if self.catalog_uri:
                configs[f"spark.sql.catalog.{self.catalog_name}.uri"] = self.catalog_uri
        else;
            configs[f"spark.sql.catalog.{self.catalog_name}.warehouse"] = self.warehouse_path
        return configs
    
    def _resolve_table(self, table:str) -> str:
        """Resolves a table reference to a fully-qualified ``catalog.namespace.table` name.
        
        * 3 parts - used as-is (caller has full control)
        * 2 parts - connection's catalog name is prepended
        * 1 part - connection's catalog name and default namespace are prepended
        """
        parts = table.split(".")
        if len(parts) >= 3:
            return table
        elif len(parts) == 2:
            return f"{self.catalog_name}.{table}"
        else:
            return f"{self.catalog_name}.{self.database}.{table}"
        
    # ----------------------------------------------------------------
    # Item 1 - branch-aware read/write
    # ----------------------------------------------------------------
    
    def read_table_into_spark(self,
                              table: str,
                              branch: Optional[str] = None,
                              partition: "PartitionSchema" = None,
                              fetch_size: Optional[int] = None
        ) -> DataFrame:
        """Reads an Iceberg table into a Spark DataFrame.

        Args:
            table (str): Table name (bare, ``namespace.table``, or fully-qualified). 
            branch (Optional[str], optional): Named branch to read from. Reads the table's main branch when omitted.
        """
        reader = self.spark.read.format("iceberg")
        if branch:
            reader = reader.option("branch", branch)
        return reader.load(self._resolve_table(table))
    
    def query_into_spark(
        self,
        query: str,
        partition: "PartitionSchema" = None,
        fetch_size: Optional[int] = None,
        **kwargs,
    ) -> DataFrame:
        """Executes a SQL query via spark.sql(), qualifying before bare table names with the catalog."""
        return self.spark.sql(query)

    def write_table_from_spark(
        self,
        df: DataFrame,
        table: str,
        mode: SaveMode = SaveMode.APPEND,
        branch: Optional[str] = None,
        max_partitions: Optional[int] = None,
        batch_size: Optional[int] = None,
        truncate: bool=False,
        **kwargs,
    ):
        """Writes a Spark DataFrame to an Iceberg table.
        Args:
            df:     DataFrame to write.
            table:  Table name (bare, ``namespace.table``, or fully-qualified).
            mode:   Write Mode. Defaults to ``SaveMode.APPEND``; Iceberg branch
                                        writes are almost always appends.
            branch: Named branch to write to. Writes to the table's main branch
                    when omitted.
            
        Note:
            Use the v1 ``DataFrameWriter`` (``df.write.format("iceberg)``) rather
            than ``DataFrameWriterV2`` (``df.writeTo``). In Spark 4.1 / Iceberg
            1.11 the v2 writer silently ignores the ``branch`` option and routes
            the write to main.            
        """
        writer = df.write.format("iceberg").mode(mode.name.lower())
        if branch:
            writer = writer.option("branch", branch)
        writer.save(self._resolve_table(table))
        
    # ----------------------------------------------------------------
    # Item 2 - branch and tag DDL
    # ----------------------------------------------------------------
    
    def _get_branch_snapshot_id(self, table:str, branch: str) -> int:
        """Returns the current tip snapshot ID of ``branch`` on ``table``.

        Raises:
            ValueError: If the branch does not exist on the table
        """
        fq = self._resolve_table(table)
        row = self.spark.sql(
            f"SELECT snapshot_id FROM {fq}.refs WHERE name = '{branch}'"
        ).first()
        if row is None:
            raise ValueError(f"Branch '{branch}' not found on table {fq}")
        return row["snapshot_id"]
    
    def ensure_branch(self, table:str, branch:str) -> None:
        """Creates ``branch`` on ``table`` if it does not already exist.
        
        Uses ``CREATE BRANCH IF NOT EXIST`` so existing branch history and
        snapshot pointers are never distributed on re-runs.
        """
        fq = self._resolve_table(table)
        self.spark.sql(f"ALTER TABLE {fq} CREATE BRANCH IF NOT EXISTS `{branch}`")
        logging.debug("Ensured branch '%s' exists on %s", branch, fq)
    
    def create_branch(self, table: str, branch: str) -> None:
        """Creates a new branch on ``table``.
        
        Raises if the branch already exists. Use :meth: `ensure_branch` for idempotent
        branch creation.        
        """
        fq = self._resolve_table(table)
        self.spark.sql(f"ALTER TABLE {fq} CREATE BRANCH `{branch}`")
        logging.debug("Create branch '%s' on '%s'", branch, fq)
    
    def apply_tag(self, table:str, tag:str, branch: Optional[str] = None) -> int:
        """Stamps `tag` on a snapshot of ``table``.
        
        If ``branch`` is given, resolves that branch's current tip from the ``.refs``
        metadata table and tags ``AS OF VERSION <snapshot_id>``.
        If ``branch`` is omitted, tags the table's current (latest committed) snapshot.
        
        Uses ``CREATE OR REPLACE TAG`` so re-running with the same tag name
        advances the tag to the latest snapshot rather than failing.

        Returns:
            The snapshot ID that was tagged.
        """
        fq = self._resolve_table(table)
        if branch:
            snapshot_id = self._get_branch_snapshot_id(table, branch)
            self.spark.sql(
                f"ALTER TABLE {fq} CREATE OR REPLACE TAG `{tag}` "
                f"AS OF VERSION {snapshot_id}"
            )
        else:
            self.spark.sql(f"ALTER TABLE {fq} CREATE OR REPLACE TAG `{tag}`")
            row = self.spark.sql(
                f"SELECT snapshot_id FROM {fq}.refs WHERE name = '{tag}'"
            ).first()
            snapshot_id = row["snapshot_id"] if row else None
        logging.info("Tagged %s snapshot %s as '%s'", fq, snapshot_id, tag)
        return snapshot_id
    
    def build_branch_at_tip(
        self, table: str, source_branch: str, target_branch: str
    ) -> int:
        """Pins ``target_branch`` to the current tip of ``source_branch``.
        
        Issues ``CREATE OR REPLACE BRANCH <target> AS OF VERSION <snapshot_id>``,
        which creates the branch if absent or resets it if already present.
        Intended for constructing ephemeral what-if scenarios without writing
        any new data files.

        Returns:
            The snapshot ID that ``target_branch`` was pinned to.
        """
        snapshot_id = self._get_branch_snapshot_id(table, source_branch)
        fq = self._resolve_table(table)
        self.spark.sql(
            f"ALTER TABLE {fq} CREATE OR REPLACE BRANCH `{target_branch}` "
            f"AS OF VERSION {snapshot_id}"
        )
        logging.info(
            "Pinned %s branch '%s' -> '%s' snapshot %s",
            fq, target_branch, source_branch, snapshot_id
        )
        return snapshot_id
    
    # ----------------------------------------------------------------
    # Item 3 - branch provenance
    # ----------------------------------------------------------------
    
    def get_branch_provenance(self, table:str, branch:str) -> Optional[str]:
        """Returns the branch whose snapshot history contains ``branch``'s current snapshot.
        
        Uses a two-phase strategy:
        
        * **Fast path** - if another branch's *current tip* equals the target snapshot,
          returns it immediately (single, ``.ref`` query)
          
        * **Slow oath** - otherwise, query ``.snapshots`` to build a ``{snapshot_id: parent_id}``
          map, then outer-loop over candidate branches and inner-walk each branch's ancestor chain
          looking for the target snapshot. The inner walk starts from the *parent* of each
          candidate's tip, since the tip itself was already tested in the fast path.

        Args:
            table: Table name (unqualified or fully-qualified)
            branch: The branch whose provenance to resolve (e.g. ``"what_if"``).

        Returns:
            The source branch name, or ``None`` if ``branch`` does not exist or
            its snapshot cannot be found in any other branch's ancestry.
        """
        fq = self._resolve_table(table)
        refs = {
            row["name"]: row["snapshot_id"]
            for row in self.spark.sql(
                f"SELECT name, snapshot_id FROM {fq}.refs"
            ).collect()
        }
        target_snapshot = refs.get(branch)
        if target_snapshot is None:
            return None
    
        # Fast path: another branch's current tip matches directly
        for candidate, tip in refs.items():
            if candidate != branch and tip == target_snapshot:
                return candidate
        
        # Slow path: walk each candidate branch's ancestor chain
        parents = {
            row["snapshot_id"]: row["parent_id"]
            for row in self.spark.sql(
                f"SELECT snapshot_id, parent_id FROM {fq}.refs"
            ).collect()
        }
        for candidate, tip in refs.items():
            if candidate == branch:
                continue
            current = parents.get(tip)    # tip itself was already checked above
            while current is not None:
                if current == target_snapshot:
                    return candidate
                current = parents.get(current)
        
        return None