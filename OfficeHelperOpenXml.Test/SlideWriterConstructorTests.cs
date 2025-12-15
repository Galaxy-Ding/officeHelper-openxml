using System;
using Xunit;
using OfficeHelperOpenXml.Core.Writers;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for SlideWriter constructor
    /// </summary>
    public class SlideWriterConstructorTests
    {
        [Fact]
        public void Constructor_WithValidRelationshipIdManager_CreatesInstance()
        {
            // Arrange
            var relationshipIdManager = new RelationshipIdManager();

            // Act
            var slideWriter = new SlideWriter(relationshipIdManager);

            // Assert
            Assert.NotNull(slideWriter);
        }

        [Fact]
        public void Constructor_WithNullRelationshipIdManager_ThrowsArgumentNullException()
        {
            // Arrange
            RelationshipIdManager? relationshipIdManager = null;

            // Act & Assert
            var exception = Assert.Throws<ArgumentNullException>(() => new SlideWriter(relationshipIdManager!));
            Assert.Equal("relationshipIdManager", exception.ParamName);
        }
    }
}
