from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy import Column
from sqlalchemy import Integer
from sqlalchemy import JSON
from sqlalchemy import String
from sqlalchemy import ForeignKey

Base = declarative_base()


class Record(Base):
    __tablename__ = "records"

    id = Column(Integer, primary_key=True)
    data = Column(JSON)
    item_id = Column(String, ForeignKey("items.uuid"), nullable=True)

    item = relationship("Item")

    def __repr__(self):
        return f"Record(id={self.id!r}, data={self.data!r})"


class Item(Base):
    __tablename__ = "items"

    id = Column(Integer, primary_key=True)
    uuid = Column(String)
    name = Column(String)

    primes = relationship("Prime", back_populates="item")

    def __repr__(self):
        return f"Item(id={self.id!r}, uuid={self.uuid!r}, name={self.name!r})"


class Prime(Base):
    __tablename__ = "primes"

    id = Column(Integer, primary_key=True)
    name = Column(String(255))
    item_id = Column(Integer, ForeignKey('items.id'), nullable=False)

    item = relationship('Item')

    def __repr__(self):
        return f"Prime(id={self.id!r}, name={self.name!r})"
