// Copyright (C) Josh Smith - January 2007
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using WPF.JoshSmith.Adorners;
using WPF.JoshSmith.Controls.Utilities;

namespace WPF.JoshSmith.ServiceProviders.UI {
    /// <summary>
    /// Manages the dragging and dropping of ListViewItems in a ListView.
    /// The ItemType type parameter indicates the type of the objects in
    /// the ListView's items source.  The ListView's ItemsSource must be
    /// set to an instance of ObservableCollection of ItemType, or an
    /// Exception will be thrown.
    /// </summary>
    /// <typeparam name="ItemType">The type of the ListView's items.</typeparam>
    public class ListViewDragDropManager<ItemType> where ItemType : class {
        #region Data

        bool canInitiateDrag;
        DragAdorner dragAdorner;
        double dragAdornerOpacity;
        int indexToSelect;
        bool isDragInProgress;
        ItemType itemUnderDragCursor;
        ListView listView;
        Point ptMouseDown;
        bool showDragAdorner;


        #endregion // Data

        #region Constructors

        /// <summary>
        /// Initializes a new instance of ListViewDragManager.
        /// </summary>
        public ListViewDragDropManager() {
            this.canInitiateDrag = false;
            this.dragAdornerOpacity = 0.7;
            this.indexToSelect = -1;
            this.showDragAdorner = true;
        }

        /// <summary>
        /// Initializes a new instance of ListViewDragManager.
        /// </summary>
        /// <param name="listView"></param>
        public ListViewDragDropManager(ListView listView) : this() {
            this.ListView = listView;
        }

        /// <summary>
        /// Initializes a new instance of ListViewDragManager.
        /// </summary>
        /// <param name="listView"></param>
        /// <param name="dragAdornerOpacity"></param>
        public ListViewDragDropManager(ListView listView, double dragAdornerOpacity) : this(listView) {
            this.DragAdornerOpacity = dragAdornerOpacity;
        }

        /// <summary>
        /// Initializes a new instance of ListViewDragManager.
        /// </summary>
        /// <param name="listView"></param>
        /// <param name="showDragAdorner"></param>
        public ListViewDragDropManager(ListView listView, bool showDragAdorner) : this(listView) {
            this.ShowDragAdorner = showDragAdorner;
        }

        #endregion // Constructors

        #region Public Interface
        /// <summary>
        /// Gets/sets the opacity of the drag adorner.  This property has no
        /// effect if ShowDragAdorner is false. The default value is 0.7
        /// </summary>
        public double DragAdornerOpacity {
            get { return this.dragAdornerOpacity; }
            set {
                if (this.IsDragInProgress)
                    throw new InvalidOperationException("Cannot set the DragAdornerOpacity property during a drag operation.");

                if (value < 0.0 || value > 1.0)
                    throw new ArgumentOutOfRangeException("DragAdornerOpacity", value, "Must be between 0 and 1.");

                this.dragAdornerOpacity = value;
            }
        }

        /// <summary>
        /// Returns true if there is currently a drag operation being managed.
        /// </summary>
        public bool IsDragInProgress {
            get { return this.isDragInProgress; }
            private set { this.isDragInProgress = value; }
        }

        /// <summary>
        /// Gets or sets optional minimal distance mouse pointer has to be dragged horizontally before it is considered a drag operation.
        /// If value is not set, <see cref="SystemParameters.MinimumHorizontalDragDistance"/> is used.
        /// Set to <see cref="Double.PositiveInfinity"/> to prevent horizontal drag
        /// (which does not make sense with standard ListView's vertical layout).
        /// </summary>
        public double? MinimumHorizontalDragDistance { get; set; }
        /// <summary>
        /// Gets or sets optional minimal distance mouse pointer has to be dragged vertically before it is considered a drag operation.
        /// If value is not set, <see cref="SystemParameters.MinimumVerticalDragDistance"/> is used.
        /// Set to <see cref="Double.PositiveInfinity"/> to prevent vertical drag.
        /// </summary>
        public double? MinimumVerticalDragDistance { get; set; }

        /// <summary>
        /// Gets or sets precondition for dragging. If set to <c>null</c> (default value) drag is always allowed.
        /// Provided <see cref="Point"/> is the mouse position relative to the list view.
        /// </summary>
        public Func<Point, bool> DraggingPrecondition { get; set; }

        /// <summary>
        /// allow x range to start dragging
        /// </summary>
        public double AllowStartX { set; get; } = -1;
        public double AllowEndX { set; get; } = -1;

        /// <summary>
        /// Gets/sets the ListView whose dragging is managed.  This property
        /// can be set to null, to prevent drag management from occuring.  If
        /// the ListView's AllowDrop property is false, it will be set to true.
        /// </summary>
        public ListView ListView {
            get { return listView; }
            set {
                if (this.IsDragInProgress)
                    throw new InvalidOperationException("Cannot set the ListView property during a drag operation.");

                if (this.listView != null) {
                    #region Unhook Events

                    this.listView.PreviewMouseLeftButtonDown -= ListView_PreviewMouseLeftButtonDown;
                    this.listView.PreviewMouseMove -= ListView_PreviewMouseMove;
                    this.listView.DragOver -= ListView_DragOver;
                    this.listView.DragLeave -= ListView_DragLeave;
                    this.listView.DragEnter -= ListView_DragEnter;
                    this.listView.Drop -= ListView_Drop;

                    #endregion // Unhook Events
                }

                this.listView = value;

                if (this.listView != null) {
                    if (!this.listView.AllowDrop)
                        this.listView.AllowDrop = true;

                    this.listView.PreviewMouseLeftButtonDown += ListView_PreviewMouseLeftButtonDown;
                    this.listView.PreviewMouseMove += ListView_PreviewMouseMove;
                    this.listView.DragOver += ListView_DragOver;
                    this.listView.DragLeave += ListView_DragLeave;
                    this.listView.DragEnter += ListView_DragEnter;
                    this.listView.Drop += ListView_Drop;
                }

            }
        }

        /// <summary>
        /// Raised when a drop occurs.  By default the dropped item will be moved
        /// to the target index.  Handle this event if relocating the dropped item
        /// requires custom behavior.  Note, if this event is handled the default
        /// item dropping logic will not occur.
        /// </summary>
        public event EventHandler<ProcessDropEventArgs<ItemType>> ProcessDrop;

        /// <summary>
        /// ドロップ完了イベント通知
        /// </summary>
        /// <remarks>ProcessDropを使用するとリストの並べ替えを自前で実装する必要があるので、並べ替えはライブラリで行った上でイベントを通知する。</remarks>
        public delegate void DropDoneHandler(int index);
        public event DropDoneHandler DropDone;

        /// <summary>
        /// Gets/sets whether a visual representation of the ListViewItem being dragged
        /// follows the mouse cursor during a drag operation.  The default value is true.
        /// </summary>
        public bool ShowDragAdorner {
            get { return this.showDragAdorner; }
            set {
                if (this.IsDragInProgress)
                    throw new InvalidOperationException("Cannot set the ShowDragAdorner property during a drag operation.");

                this.showDragAdorner = value;
            }
        }

        // check if the target listview item is valid(can move)
        public delegate bool IsValidItemDelegate(int index);
        public IsValidItemDelegate IsValidItem;

        #endregion // Public Interface

        #region Event Handling Methods
        void ListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (this.IsMouseOverScrollbar) {
                // 4/13/2007 - Set the flag to false when cursor is over scrollbar.
                this.canInitiateDrag = false;
                return;
            }

            int index = this.IndexUnderDragCursor;
            this.canInitiateDrag = index > -1;

            if (this.canInitiateDrag) {
                // Remember the location and index of the ListViewItem the user clicked on for later.
                this.ptMouseDown = MouseUtilities.GetMousePosition(this.listView);
                this.indexToSelect = index;
            } else {
                this.ptMouseDown = new Point(-10000, -10000);
                this.indexToSelect = -1;
            }
        }

        void ListView_PreviewMouseMove(object sender, MouseEventArgs e) {
            if (!this.CanStartDragOperation)
                return;

            // Select the item the user clicked on.
            if (this.listView.SelectedIndex != this.indexToSelect)
                this.listView.SelectedIndex = this.indexToSelect;

            // If the item at the selected index is null, there's nothing
            // we can do, so just return;
            if (this.listView.SelectedItem == null)
                return;

            if (null != this.IsValidItem && !this.IsValidItem(this.listView.SelectedIndex)) {
                return;
            }

            ListViewItem itemToDrag = this.GetListViewItem(this.listView.SelectedIndex);
            if (itemToDrag == null)
                return;

            AdornerLayer adornerLayer = this.ShowDragAdornerResolved ? this.InitializeAdornerLayer(itemToDrag) : null;

            this.InitializeDragOperation(itemToDrag);
            this.PerformDragOperation();
            this.FinishDragOperation(itemToDrag, adornerLayer);
        }

        void ListView_DragOver(object sender, DragEventArgs e) {
            e.Effects = DragDropEffects.Move;

            if (this.ShowDragAdornerResolved)
                this.UpdateDragAdornerLocation();

            // Update the item which is known to be currently under the drag cursor.
            int index = this.IndexUnderDragCursor;

            if (this.IsValidItem != null && !this.IsValidItem(index)) {
                e.Effects = DragDropEffects.None;
            } else {
                e.Effects = DragDropEffects.Move;
                this.ItemUnderDragCursor = index < 0 ? null : this.ListView.Items[index] as ItemType;
            }
        }

        void ListView_DragLeave(object sender, DragEventArgs e) {
            if (!this.IsMouseOver(this.listView)) {
                if (this.ItemUnderDragCursor != null)
                    this.ItemUnderDragCursor = null;

                if (this.dragAdorner != null)
                    this.dragAdorner.Visibility = Visibility.Collapsed;
            }
        }

        void ListView_DragEnter(object sender, DragEventArgs e) {
            if (this.dragAdorner != null && this.dragAdorner.Visibility != Visibility.Visible) {
                // Update the location of the adorner and then show it.
                this.UpdateDragAdornerLocation();
                this.dragAdorner.Visibility = Visibility.Visible;
            }
        }

        void ListView_Drop(object sender, DragEventArgs e) {
            if (this.ItemUnderDragCursor != null)
                this.ItemUnderDragCursor = null;

            e.Effects = DragDropEffects.None;

            if (!e.Data.GetDataPresent(typeof(ItemType)))
                return;

            // Get the data object which was dropped.
            if (!(e.Data.GetData(typeof(ItemType)) is ItemType data))
                return;

            // Get the ObservableCollection<ItemType> which contains the dropped data object.
            if (!(this.listView.ItemsSource is ObservableCollection<ItemType> itemsSource))
                throw new Exception(
                    "A ListView managed by ListViewDragManager must have its ItemsSource set to an ObservableCollection<ItemType>.");

            int oldIndex = itemsSource.IndexOf(data);
            int newIndex = this.IndexUnderDragCursor;

            if (null != IsValidItem && !IsValidItem(newIndex)) {
                // 先頭のカテゴリの上には移動させない
                if (0 == newIndex) {
                    newIndex = 1;
                }
            }

            if (newIndex < 0) {
                // The drag started somewhere else, and our ListView is empty
                // so make the new item the first in the list.
                if (itemsSource.Count == 0)
                    newIndex = 0;

                // The drag started somewhere else, but our ListView has items
                // so make the new item the last in the list.
                else if (oldIndex < 0)
                    newIndex = itemsSource.Count;

                // The user is trying to drop an item from our ListView into
                // our ListView, but the mouse is not over an item, so don't
                // let them drop it.
                else
                    return;
            }

            // Dropping an item back onto itself is not considered an actual 'drop'.
            if (oldIndex == newIndex)
                return;

            if (this.ProcessDrop != null) {
                // Let the client code process the drop.
                ProcessDropEventArgs<ItemType> args = new ProcessDropEventArgs<ItemType>(itemsSource, data, oldIndex, newIndex, e.AllowedEffects);
                this.ProcessDrop(this, args);
                e.Effects = args.Effects;
            } else {
                // Move the dragged data object from it's original index to the
                // new index (according to where the mouse cursor is).  If it was
                // not previously in the ListBox, then insert the item.
                if (oldIndex > -1)
                    itemsSource.Move(oldIndex, newIndex);
                else
                    itemsSource.Insert(newIndex, data);

                // Set the Effects property so that the call to DoDragDrop will return 'Move'.
                e.Effects = DragDropEffects.Move;
                DropDone?.Invoke(newIndex);
            }
        }
        #endregion // Event Handling Methods

        #region Private Helpers

        bool CanStartDragOperation {
            get {
                if (Mouse.LeftButton != MouseButtonState.Pressed)
                    return false;

                if (!this.canInitiateDrag)
                    return false;

                if (this.indexToSelect == -1)
                    return false;

                if (false.Equals(this.DraggingPrecondition?.Invoke(this.ptMouseDown)))
                    return false;

                if (0 <= AllowStartX && 0 <= AllowEndX) {
                    if (this.ptMouseDown.X < AllowStartX || AllowEndX < this.ptMouseDown.X) {
                        return false;
                    }
                }

                // System.Diagnostics.Debug.WriteLine(this.ptMouseDown.X + ":" + this.ptMouseDown.Y);


                return this.HasCursorLeftDragThreshold;
            }
        }

        void FinishDragOperation(ListViewItem draggedItem, AdornerLayer adornerLayer) {
            // Let the ListViewItem know that it is not being dragged anymore.
            ListViewItemDragState.SetIsBeingDragged(draggedItem, false);

            this.IsDragInProgress = false;

            if (this.ItemUnderDragCursor != null)
                this.ItemUnderDragCursor = null;

            // Remove the drag adorner from the adorner layer.
            if (adornerLayer != null) {
                adornerLayer.Remove(this.dragAdorner);
                this.dragAdorner = null;
            }
        }

        ListViewItem GetListViewItem(int index) {
            if (this.listView.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                return null;

            return this.listView.ItemContainerGenerator.ContainerFromIndex(index) as ListViewItem;
        }

        ListViewItem GetListViewItem(ItemType dataItem) {
            if (this.listView.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                return null;

            return this.listView.ItemContainerGenerator.ContainerFromItem(dataItem) as ListViewItem;
        }

        bool HasCursorLeftDragThreshold {
            get {
                if (this.indexToSelect < 0)
                    return false;

                ListViewItem item = this.GetListViewItem(this.indexToSelect);
                Rect bounds = VisualTreeHelper.GetDescendantBounds(item);
                Point ptInItem = this.listView.TranslatePoint(this.ptMouseDown, item);

                // In case the cursor is at the very top or bottom of the ListViewItem
                // we want to make the vertical threshold very small so that dragging
                // over an adjacent item does not select it.
                double topOffset = Math.Abs(ptInItem.Y);
                double btmOffset = Math.Abs(bounds.Height - ptInItem.Y);
                double vertOffset = Math.Min(topOffset, btmOffset);

                double width = (this.MinimumHorizontalDragDistance ?? SystemParameters.MinimumHorizontalDragDistance) * 2;
                double height = Math.Min(this.MinimumVerticalDragDistance ?? SystemParameters.MinimumVerticalDragDistance, vertOffset) * 2;
                Size szThreshold = new Size(width, height);

                Rect rect = new Rect(this.ptMouseDown, szThreshold);
                rect.Offset(szThreshold.Width / -2, szThreshold.Height / -2);
                Point ptInListView = MouseUtilities.GetMousePosition(this.listView);
                return !rect.Contains(ptInListView);
            }
        }

        /// <summary>
        /// Returns the index of the ListViewItem underneath the
        /// drag cursor, or -1 if the cursor is not over an item.
        /// </summary>
        int IndexUnderDragCursor {
            get {
                int index = -1;
                for (int i = 0; i < this.listView.Items.Count; ++i) {
                    ListViewItem item = this.GetListViewItem(i);
                    if (item == null) {
                        continue;
                    }
                    if (this.IsMouseOver(item)) {
                        index = i;
                        break;
                    }
                }
                return index;
            }
        }

        AdornerLayer InitializeAdornerLayer(ListViewItem itemToDrag) {
            // Create a brush which will paint the ListViewItem onto
            // a visual in the adorner layer.
            VisualBrush brush = new VisualBrush(itemToDrag);

            // Create an element which displays the source item while it is dragged.
            this.dragAdorner = new DragAdorner(this.listView, itemToDrag.RenderSize, brush) {

                // Set the drag adorner's opacity.
                Opacity = this.DragAdornerOpacity
            };

            AdornerLayer layer = AdornerLayer.GetAdornerLayer(this.listView);
            layer.Add(dragAdorner);

            // Save the location of the cursor when the left mouse button was pressed.
            this.ptMouseDown = MouseUtilities.GetMousePosition(this.listView);

            return layer;
        }

        void InitializeDragOperation(ListViewItem itemToDrag) {
            // Set some flags used during the drag operation.
            this.IsDragInProgress = true;
            this.canInitiateDrag = false;

            // Let the ListViewItem know that it is being dragged.
            ListViewItemDragState.SetIsBeingDragged(itemToDrag, true);
        }

        bool IsMouseOver(Visual target) {
            // We need to use MouseUtilities to figure out the cursor
            // coordinates because, during a drag-drop operation, the WPF
            // mechanisms for getting the coordinates behave strangely.

            Rect bounds = VisualTreeHelper.GetDescendantBounds(target);
            Point mousePos = MouseUtilities.GetMousePosition(target);
            return bounds.Contains(mousePos);
        }

        /// <summary>
        /// Returns true if the mouse cursor is over a scrollbar in the ListView.
        /// </summary>
        bool IsMouseOverScrollbar {
            get {
                Point ptMouse = MouseUtilities.GetMousePosition(this.listView);
                HitTestResult res = VisualTreeHelper.HitTest(this.listView, ptMouse);
                if (res == null)
                    return false;

                DependencyObject depObj = res.VisualHit;
                while (depObj != null) {
                    if (depObj is ScrollBar)
                        return true;

                    // VisualTreeHelper works with objects of type Visual or Visual3D.
                    // If the current object is not derived from Visual or Visual3D,
                    // then use the LogicalTreeHelper to find the parent element.
                    if (depObj is Visual || depObj is System.Windows.Media.Media3D.Visual3D)
                        depObj = VisualTreeHelper.GetParent(depObj);
                    else
                        depObj = LogicalTreeHelper.GetParent(depObj);
                }

                return false;
            }
        }

        ItemType ItemUnderDragCursor {
            get { return this.itemUnderDragCursor; }
            set {
                if (this.itemUnderDragCursor == value)
                    return;

                // The first pass handles the previous item under the cursor.
                // The second pass handles the new one.
                for (int i = 0; i < 2; ++i) {
                    if (i == 1)
                        this.itemUnderDragCursor = value;

                    if (this.itemUnderDragCursor != null) {
                        ListViewItem listViewItem = this.GetListViewItem(this.itemUnderDragCursor);
                        if (listViewItem != null)
                            ListViewItemDragState.SetIsUnderDragCursor(listViewItem, i == 1);
                    }
                }
            }
        }

        void PerformDragOperation() {
            ItemType selectedItem = this.listView.SelectedItem as ItemType;
            DragDropEffects allowedEffects = DragDropEffects.Move | DragDropEffects.Move | DragDropEffects.Link;
            if (DragDrop.DoDragDrop(this.listView, selectedItem, allowedEffects) != DragDropEffects.None) {
                // The item was dropped into a new location,
                // so make it the new selected item.
                this.listView.SelectedItem = selectedItem;
            }
        }

        bool ShowDragAdornerResolved {
            get { return this.ShowDragAdorner && this.DragAdornerOpacity > 0.0; }
        }

        void UpdateDragAdornerLocation() {
            if (this.dragAdorner != null) {
                Point ptCursor = MouseUtilities.GetMousePosition(this.ListView);

                double left = ptCursor.X - this.ptMouseDown.X;

                // 4/13/2007 - Made the top offset relative to the item being dragged.
                ListViewItem itemBeingDragged = this.GetListViewItem(this.indexToSelect);
                Point itemLoc = itemBeingDragged.TranslatePoint(new Point(0, 0), this.ListView);
                double top = itemLoc.Y + ptCursor.Y - this.ptMouseDown.Y;

                this.dragAdorner.SetOffsets(left, top);
            }
        }

        #endregion // Private Helpers
    }
}